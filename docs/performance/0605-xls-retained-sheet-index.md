# 0605: one validated scan for a whole XLS worksheet, and a frozen design for the retained sheet index

Status: retained for part (1), the whole-sheet walk; part (2), the snapshot-scoped
retained sheet index, stops at a frozen design and is **not** implemented.
`performance_claim: none` — this record carries deterministic counts, callgrind
isolation pairs, native hardware counters and paired medians in both directions
with a measured noise floor, not a claim-registry entry.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item **XLS-3** (rank 18) of change
[0587](0587-remaining-opportunity-survey.md), whose two halves the survey
deliberately separated: *"even without the index, a public
`SourceBackedWorksheet::cells()` turns N scans into one and is the missing
all-cells selector"*, against an index that *"must be weighted, bounded by
`max_worksheet_scan_records` and evictable, and `litchi-xls` has no such cache
today — a frozen design record"*. Part (1) is the first half. Part (2) is the
second, and the gate this record was given for implementing it — *value-identical
and the retention accounting in place* — is not met, so it is written down and
left unbuilt.

The sentence this change closes is change [0568](0568-xls-worksheet-window.md)'s
third limitation, carried unaddressed for 37 records: *"the source-backed API has
no whole-sheet iterator, so a 'full cell scan' is N independent one-cell queries,
each re-scanning its sheet; only the per-query ratio is meaningful."*

## What was changed

Three files: `crates/litchi-xls/src/workbook/source.rs`,
`crates/litchi-xls/tests/source_backed.rs`, and the standalone attribution
harness `tools/perf-baseline/src/bin/xls_source_attribution.rs` with its README
entry.

### One frame loop, two sinks

`query_cell` framed the selected worksheet's substream to its EOF, parsed every
cell record, and handed each one to `process_cell`, which validated its XF index,
compared its `(row, col)` against the target and discarded all but the match. The
loop body is now `scan_worksheet`, generic over a private `CellSink`:

```rust
trait CellSink {
    fn accept(&mut self, record: &CellRecord, scan: &ScanContext<'_>,
              strings: &mut SharedStringResolver<'_>) -> Result<()>;
    fn defers_string_formula(&self, record: &CellRecord) -> bool;
}
```

`TargetCell` is `process_cell` moved behind that trait, unchanged line for line.
`VisitCells<F>` reports every stored cell to a caller's visitor. The loop itself
— the BOF check, the EOF check, the pending-`FORMULA` protocol, the six cell
record kinds, `MulRk`, `MulBlank`, the `skip_payload` default, the leading and
trailing freshness fence — is the text that was in `query_cell`, moved once.

`defers_string_formula` is the only place the two sinks take different branches,
and it is exactly the branch `query_cell` already had: a string-valued `FORMULA`
is held back until its `STRING` result arrives **only for the cell the caller
asked for**, because every other cell's value was discarded. A whole-sheet walk
reports all of them, so it holds back all of them.

`ScanContext` gathers the three per-scan constants (`owner`, `formatting`,
`execution`) that `process_cell` took as separate arguments. This is not
cosmetic: see *A regression that had to be designed away*.

### The public walk

```rust
impl SourceBackedWorksheet {
    pub fn visit_cells<F>(&self, visitor: F) -> Result<()>
        where F: FnMut(SourceBackedCell) -> Result<()>;
    pub fn visit_cells_with_execution<F>(&self, execution: &ExecutionContext, visitor: F)
        -> Result<()>
        where F: FnMut(SourceBackedCell) -> Result<()>;
}
```

A visitor rather than an `Iterator` or a `Vec`, for one reason: bounded
resources. A `cells() -> Result<Vec<SourceBackedCell>>` retains every cell of the
sheet — 38,950 of them on `54016.xls`, string payloads included — with no limit
in `SourceBackedLimits` that bounds it, and `max_text_cells` is documented as the
text projection's budget, not a general one. The visitor retains nothing the
caller does not choose to retain, and a caller that wants a `Vec` writes its own
bound. An `Iterator` is not available without either self-referential borrows of
the scan state or retaining the whole sheet first, which is the same problem.

The visitor **cannot stop the scan**. `visit_cells` always runs to the
worksheet's EOF, because `query_cell` does: an early exit would be change 0574's
opportunity 6, which ADR 0005's mandatory-validation clause rejects and which the
survey's own "not opportunities" list names. A visitor that returns an error does
end the scan, and that error is returned unchanged — the same way any other
error ends it.

### Three scenarios the harness could not express

`xls_source_attribution` supported `open`, `list` and `one-cell`. The survey
listed the consequence as a measurement blocker: *"no source-backed XLS full-text
or all-cells selector exists … so XLS-3, 4, 8 and 10 cannot be sized today."* It
now also supports:

- `--operation all-cells`, with `--all-cells-strategy scan` (the walk) or
  `per-cell` (one `cell_value_by_index` per position, bounded by
  `--per-cell-limit`, default 64). These are the paired legs of the scenario
  **inside one binary and one build**, because the before leg has no walk at all;
  `per-cell` uses only API that predates this change.
- `--operation full-text`, which records a typed refusal as an outcome instead of
  failing the run — three worksheets of the flagship fixture are refused by this
  reader, and holding that refusal identical across legs is the point.
- `--operation second-cell`, two queries on one worksheet: the scenario a
  retained index would serve, measured so that part (2)'s design can state what
  it would save.

The all-cells oracle runs on **every** invocation, outside the timed region: it
walks the sheet, re-reads a bounded prefix of the reported positions one
selected-cell query at a time, and fails the run if the two projections differ.
Every all-cells measurement in this record is therefore also a differential.

## Why it is sound

**No refusal moves.** Every check `query_cell` took, `scan_worksheet` takes, in
the same order, against the same bytes, returning the same
`SourceBackedError`. The one textual addition is a `try_reserve` before the
`Vec::push` of a `FORMULA` `STRING` continuation, which `scan_text_sheet` already
had and `query_cell` did not; it can only convert an allocator abort into
`SourceBackedError::Allocation`, which is the direction GOAL rule 12 asks for and
never the reverse. The counters below show the conversion is unreachable on every
fixture measured: reads, bytes and source observations are identical to the byte
on both legs.

**Error identity is proved, not asserted.**
`the_whole_sheet_walk_refuses_what_a_selected_cell_query_refuses` compares
`Display` strings, not variants, on a synthetic truncated `Number` record and on
the three real flagship worksheets whose shared-formula metadata this reader
declines. All four refusals are string-equal between the walk and a query.

**The freshness bracket is unchanged.** `visit_cells` takes the leading
cancellation check and `ensure_current()` that `query_cell` takes, and
`scan_worksheet` takes the trailing pair at EOF that `finish_query` took. The
walk of `WithCustomViews.xls` sheet 0 and a one-cell query on the same sheet
observe the source the same number of times for the scan itself; the walk's extra
observations are the three per resolved shared string that
`resolve_shared_string` has always taken, one set per string rather than one set
per query.

**ADR 0005.** The walk adds no cache, retains nothing between calls, and loads no
payload eagerly: it is the same lazy scan `query_cell` performs, exposed once
instead of once per cell. ADR 0003's "cheap-to-share `Send + Sync` snapshots" is
untouched — `SourceBackedWorksheet` is still two words and still `Clone`.
ADR 0006 is not engaged: nothing is written and no validation was moved or
weakened.

**No new `unsafe`, no new dependency, no public leakage.** `CellSink`,
`ScanContext`, `TargetCell` and `VisitCells` are private. The only public surface
added is the two methods and the `FnMut(SourceBackedCell) -> Result<()>` bound
over types that were already public.

### A regression that had to be designed away

The first working version of the shared loop cost **+1.37% instructions and
+1.36% cycles** on the `54016.xls` one-cell query — the heaviest query in the
corpus and the one change 0595 had just made leaner. Callgrind located it: with
`accept` taking `owner`, `formatting`, `execution` and `strings` as four separate
arguments, `query_cell`'s own instruction count rose by 453,931 (about 12 per
frame over 37,929 frames) and `next_frame` and `read_payload` each gained exactly
one instruction per call, the signature of register pressure in the loop rather
than of new work. Two attempts failed to recover it: `#[inline]` on both sink
methods (18,856,289 → 18,817,353 Ir, still +1.16%) and restoring the original
nested branch shape around the string-formula case (18,817,347 Ir, unchanged).
Folding the three constants into `ScanContext` — two pointers to the sink instead
of four — recovered it to **+0.011%** (18,603,993 against 18,601,973). The three
trial profiles are retained in the packet. The lesson is worth the two paragraphs
it costs: sharing a hot loop across sinks is free only if the sink call's
argument count is.

## Measured

Host, toolchain, binary and fixture identities: `results/change-0605/README.md`
and `environment.json`. Every table here is reproduced by
`results/change-0605/analyze.py` over that directory alone.

### The control: not one byte of I/O moved on the existing operations

Three fixtures × two in-memory source modes × `open`, `list`, `one-cell`, before
against after. Every cell is identical in reads, read bytes, source-version
observations, `len` calls and seeks, and every sample within a cell carries
identical counters.

| Fixture | Operation | reads | bytes | observations |
| --- | --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | open / list | 53 | 565,201 | 29 |
| | one-cell | 61 | 603,115 | 42 |
| `WithCustomViews.xls` | open / list | 16 | 110,242 | 22 |
| | one-cell | 17 | 110,656 | 25 |
| `54016.xls` | open / list | 40 | 317,171 | 25 |
| | one-cell | 65 | 932,993 | 43 |

(Both legs; the before and after columns are equal, so one column is shown.)

### The walk against the per-cell reading it replaces

`owned-readat`, the worksheet of each fixture that actually stores cells.
`per-cell` reads a bounded prefix of the same positions, because each position
costs a complete validated scan of the substream.

| Fixture, sheet | Leg | cells | reads | bytes | observations | p50 |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `54016`, 0 | walk | 38,950 | 16,145 | 1,256,139 | 64,307 | 20.71 ms |
| | per-cell ×8 | 8 | 248 | 5,243,842 | 208 | 5.30 ms |
| | per-cell ×64 | 64 | 1,690 | 39,731,462 | 1,440 | 40.66 ms |
| | per-cell ×256 | 256 | 6,529 | 157,972,311 | 5,244 | 161.99 ms |
| `WithCustomViews`, 0 | walk | 3,325 | 896 | 257,066 | 3,501 | 1.69 ms |
| | per-cell ×8 | 8 | 73 | 434,688 | 105 | 0.39 ms |
| | per-cell ×64 | 64 | 474 | 2,705,634 | 701 | 2.90 ms |
| | per-cell ×256 | 256 | 1,910 | 10,499,248 | 2,987 | 11.67 ms |
| flagship, 11 | walk | 347 | 108 | 600,726 | 226 | 0.198 ms |
| | per-cell ×8 | 8 | 124 | 846,312 | 136 | 0.174 ms |
| | per-cell ×64 | 64 | 580 | 2,813,660 | 728 | 0.793 ms |
| | per-cell ×256 | 256 | 2,141 | 9,558,820 | 2,748 | 2.893 ms |

Read **bytes** is the cleanest column, because it does not depend on the window
schedule: reading 256 of `54016`'s cells one at a time moves 157,972,311 bytes;
walking all 38,950 moves 1,256,139. That is 152× the cells for 1/126 the bytes,
and **32.3 bytes per cell against 617,079.3**, a factor of 19,134. On
`WithCustomViews` the factor is 530×, on the flagship's densest worksheet 22×.

The per-cell leg is linear in the positions read — 8 : 64 : 256 is
5,243,842 : 39,731,462 : 157,972,311, within 5% of 1 : 7.58 : 30.1 — so the whole
sheet extrapolates to 38,950/256 × 161.99 ms ≈ **24.6 s** against the walk's
20.71 ms, about 1,190×. `WithCustomViews` extrapolates to 151.6 ms against
1.69 ms, about 90×; the flagship's sheet 11 to 3.9 ms against 0.198 ms, about
20×. The extrapolation is **modelled**; the three measured points it rests on are
not.

The crossover, from the measured points: the walk of all 347 flagship cells costs
about what 9 per-cell queries cost, the walk of all 3,325 `WithCustomViews` cells
about what 35 cost, and the walk of all 38,950 `54016` cells about what 31 cost.
A caller who wants more than about three dozen cells of a worksheet is already
better off walking it.

### Instructions per operation, callgrind isolation pairs

`(large − small) / (large samples − small samples)`, change 0574's method.
Callgrind's counts are deterministic, so these differences are signal at a
resolution wall clock cannot reach.

| Fixture | Operation | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| flagship | open | 2,326,978 | 2,327,785 | +0.035% |
| flagship | one-cell | 2,554,707 | 2,556,593 | +0.074% |
| cv | open | 808,140 | 809,320 | +0.146% |
| cv | one-cell | 827,903 | 828,733 | +0.100% |
| `54016` | open | 5,382,679 | 5,383,543 | +0.016% |
| `54016` | one-cell | 18,601,973 | 18,603,996 | +0.011% |

The walk, after leg only, against 64 per-cell queries on the same worksheet:

| Fixture, sheet | Leg | Ir per operation | Ir per cell |
| --- | --- | ---: | ---: |
| flagship, 1 (87 cells) | walk | 3,598,822 | 10,371 |
| | per-cell ×64 | 15,400,828 | 240,638 |
| cv, 0 (3,325 cells) | walk | 31,365,208 | 9,433 |
| | per-cell ×64 | 57,584,245 | 899,754 |
| `54016`, 0 (38,950 cells) | walk | 356,255,103 | 9,146 |
| | per-cell ×64 | 853,156,078 | 13,330,564 |

On `54016.xls` the walk reports 609× the cells for 0.42× the instructions:
**9,146 instructions per cell against 13,330,564**, a factor of 1,457.

### Cycles and IPC, native `perf stat`, same isolation method

Single-shot counters, so unlike the instruction counts they carry the host's
variance; they are here because change 0579 showed instruction share mis-ranks
pointer-chase work.

| Fixture | Operation | before cycles | after cycles | delta | IPC before | IPC after |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| flagship | open | 345,046 | 342,917 | −0.62% | 3.41 | 3.44 |
| flagship | one-cell | 397,486 | 386,806 | −2.69% | 3.38 | 3.47 |
| cv | open | 126,433 | 128,466 | +1.61% | 4.05 | 3.99 |
| cv | one-cell | 131,707 | 131,967 | +0.20% | 4.04 | 4.04 |
| `54016` | open | 867,488 | 864,836 | −0.31% | 4.65 | 4.66 |
| `54016` | one-cell | 3,672,764 | 3,702,486 | +0.81% | 4.37 | 4.34 |

Spread −2.69% to +1.61% against an instruction spread of +0.011% to +0.146%:
the cycle column is the host, not the change.

The walk against per-cell in cycles: `54016` 92,015,877 against 182,832,246;
cv 7,676,448 against 13,022,164; flagship 659,908 against 2,768,559.

### Wall clock, paired, A1 B1 B2 A2

20 warmups and 60 samples per round for `open`, `list`, `one-cell`; 5 and 30 for
the heavier new scenarios. `a1`/`a2` are the before binary, `b1`/`b2` the after
binary, so `a1` against `a2` is the **A/A floor** and `b1` against `b2` the B/B
floor, both in the same window.

**The floor.** Over the 18 cells that exist on both legs: A/A **p50 0.71%, mean
0.73%, max 2.14%** — a quiet window by this host's standards (change 0568
measured p50 4.10%, p99 13.70%). Over all 42 after-leg cells the B/B floor is
p50 1.28%, mean 3.57%, **max 46.51%**, and that maximum matters: see below.

**The 36 paired comparisons** (18 cells × two directions) run −2.66% to +7.19%,
9 of them improving. **Two exceed the +5% review trigger**, and both are
reported here rather than folded into a mean:

| Cell | b1/a1 | b2/a2 | B/B floor in the same cell |
| --- | ---: | ---: | ---: |
| `WithCustomViews.xls` `file-source` open | +1.11% | **+7.19%** | 6.67% |
| `WithCustomViews.xls` `file-source` list | +1.38% | **+6.75%** | 6.14% |

Both are on the fixture with the smallest absolute times in the matrix — 34.7 µs
— where a 2 µs shift is 6%. Both trigger in one direction only. Both sit inside
the after binary's own floor in that same cell. And both are `open` and `list`,
which **never enter the code this change touched**: the worksheet frame loop is
reached only by a cell query, and their instruction counts moved by +0.146% and
0%. The honest reading is that this cell was noisy in that window, not that the
change cost 7%; the honest report is that the trigger fired and this is why it is
not being claimed as a regression.

**The heavy scenarios are not claimed in wall clock at all.** The B/B floor on
`54016` `owned-readat` all-cells-scan is **46.51%** (20.46 ms against 29.97 ms,
same binary) and on its full text 23.97%: 30 samples of a 20-30 ms operation on a
host with eight measurement agents active is not enough. Their evidence in this
record is the deterministic counters and instruction counts, which do not have
that problem. The walk-against-per-cell ratios survive it anyway — 160.0 ms
against 20.5-30.0 ms is outside any 46% floor — but the ratio is what is claimed,
not either absolute.

**The new scenarios' medians, after leg, `owned-readat`, for the record:**
second-cell 1.469 ms / 29.5 µs / 89.3 µs (`54016` / cv / flagship); full text
24.97 ms / 1.398 ms / 0.522 ms; walk 20.46 ms / 1.703 ms / 0.196 ms.

## Correctness evidence

Eight tests added to `crates/litchi-xls/tests/source_backed.rs`:

| Test | What it proves |
| --- | --- |
| `the_whole_sheet_walk_agrees_with_selected_cell_queries` | Over the bounding rectangle of the walk's report, on three worksheets of two fixtures, every position agrees with an independent `cell()` query — including the positions the walk did **not** report, which must come back absent. |
| `..._over_formulas` / `..._reports_a_string_valued_formula_with_its_string_result` | The one branch where the walk and a query differ: a string-valued `FORMULA` with a multi-`CONTINUE` `STRING` result is reported by the walk with the value the query returns. |
| `the_whole_sheet_walk_reports_shared_strings` | The `LabelSst` resolve path is exercised. |
| `the_whole_sheet_walk_costs_one_scan_not_one_per_cell` | The per-cell path's reads and bytes are linear in the positions read; the walk of 3,325 cells costs fewer reads, bytes and observations than 256 per-cell queries. |
| `the_whole_sheet_walk_refuses_what_a_selected_cell_query_refuses` | Four refusals — one synthetic, three real — are `Display`-equal between the walk and a query. |
| `the_whole_sheet_walk_returns_a_visitor_error_unchanged` | A visitor error ends the scan at the first cell and is returned by value. |
| `the_whole_sheet_walk_observes_a_source_change` | A source mutation is refused as `SourceChanged`. |
| `the_whole_sheet_walk_honours_cancellation` | A cancelled `ExecutionContext` is refused. |

Two harness tests added: the new operations and knobs parse and default
explicitly, and the position-keyed digest ignores report order and lets the last
record for a position win, as a query does.

### The corpus differential

The brief's oracle is *"identical cell values, errors, reads, bytes and
observations for every query on every XLS fixture"*, not only the three the
timing matrix uses. `capture_corpus.sh` sweeps **all 126 `.xls` fixtures under
`test-data`** — three operations × two worksheet indices, 756 cells — on both
legs, recording each cell's logical counters, the harness's source and eager
semantic projections, and, where the reader declines a fixture, its typed
refusal. `diff corpus-before.txt corpus-after.txt` is **empty**: 756 identical
lines including **99 refusals with identical messages**, among them 23
`WorksheetNotFound`, 18 `EncryptedUnsupported`, 12 corrupted-FAT refusals, and
the malformed-`XFCRC`, short-`SST`, invalid-font and reserved-formula-bit
families.

### The walk against queries, corpus-wide

`capture_walk_differential.sh` runs `--operation all-cells` over the same 126
fixtures at worksheet indices 0, 1 and 2. Every such run builds its oracle by
walking the worksheet and then re-reading up to 128 of the reported positions one
`cell_value_by_index` call at a time, failing the run on any disagreement, so the
sweep is a corpus-wide differential between the walk and the query it is meant to
replace. Result: **291 worksheets walked, 90,543 cells reported, 145 worksheets
storing at least one, and zero disagreements.** The 89 cells that did not walk
are 56 worksheet indices a fixture does not have, 9 encrypted workbooks, and 24
fixtures this reader declines for the structural reasons the corpus sweep also
records — every one of them refused with the message the query path produces.
The widest walks are `54016.xls` sheet 0 at 38,950 cells, `45365-2.xls` at
16,598, `external_name.xls` sheet 1 at 5,344 and `FormulaEvalTestData.xls` at
4,627.

### What the change could not have touched

The text projection — `scan_text_sheet`, `write_text_sheet`,
`collect_source_cell`, `decode_source_cell` and `SourceTextSheet` — and all of
`crates/litchi-xls/src/records.rs`, where the SST scan lives, are **byte-identical
between the legs**: `git diff` touches four regions of `source.rs`
(`impl SourceBackedWorksheet`, the new sinks after `validate_worksheet_bof`,
`query_cell`/`finish_query`, and the removal of `process_cell`) and no other
file under `crates/`. Change 0576's and 0595's SST differential therefore cannot
have moved, and the flagship's refused text extraction is refused by unchanged
code — which the after leg also records directly, as
`refused:source XLS parse error: Invalid record 0x0006: shared Formula metadata
requires a leading PtgExp token`.

Gates: `results/change-0605/gates.txt`.

## Validation preserved

- Every check `query_cell` took is taken, in order, by `scan_worksheet`.
- The `try_reserve` added before a `FORMULA` `STRING` continuation push is
  strictly additive and unreachable on every fixture measured.
- The walk always scans to EOF; it cannot be used to skip validation.
- `visit_cells` reports a position stored twice twice, in stream order, because
  that is what the worksheet contains; a query reports the last such record, and
  the differential oracle folds the walk the same way before comparing.
- Nothing is retained between calls, so two walks cost two scans and the
  bounded-memory contract is exactly what it was.

## Part (2): frozen design for the snapshot-scoped retained sheet index

This section is the record. **Nothing in it is implemented**, and the reason is
stated at the end rather than left to be inferred.

### The shape

One index per `(snapshot, worksheet)`, built by the first scan of that worksheet
that reaches its EOF without returning an error, and only by such a scan:

```text
SheetIndex {
    expected_version: SourceVersion,   // the snapshot's, captured at build
    slots: Box<[CellSlot]>,            // sorted by (row, column)
}
CellSlot {
    stream_offset: u64,   // the framed record's header, in workbook-stream space
    row: u16, column: u16,
    kind: u16,            // the BIFF record kind, so the reader parses the right one
    xf: u16,              // the XF index the first scan already validated
    ordinal: u16,         // which sub-cell of a MulRk / MulBlank, else 0
}
```

Eighteen bytes, padded to 24 by alignment, or 16 packed with `stream_offset`
narrowed to `u32` (a worksheet substream cannot exceed
`max_worksheet_scan_bytes`, but the offset is absolute, so `u64` is the honest
width). Sorted slots need no separate position table: a second query binary
searches them.

There is no "validated to EOF" flag. The index's existence **is** the flag,
because it is built only by a scan that reached EOF, and that is what makes the
refusal analysis below hold.

### Lock discipline

`SourceInner` is already `Send + Sync` behind an `Arc` and holds no lock over
cell data. The natural shape is the one changes
[0193](changes/0193-ods-source-edit-protection-cache.md),
[0195](changes/0195-ods-source-content-layout-cache.md),
[0198](changes/0198-ods-content-layout-topology-cache.md) and
[0592](0592-docx-lazy-paragraph-index.md) already use: a per-sheet
`OnceLock<Option<Arc<SheetIndex>>>`, published once, with a concurrent first use
allowed to compute twice and publish one result. That shape is **not admissible
here**, and that is the crux of this design; see *Why it is not implemented*.

An admissible shape is a `Mutex<WeightedSheetCache>` on `SourceInner` holding a
byte budget, an eviction order, and `Arc` handles so that a walk in progress pins
its index against eviction. Every lock acquisition would sit outside the scan, on
the lookup and the publish, never around I/O.

### What a second query does

Check `expected_version` against the snapshot's; on a mismatch, discard and scan.
Otherwise binary search `slots` for `(row, column)`; on a miss, return `Ok(None)`
without reading. On a hit, construct the cursor, seek to `stream_offset`, read
that one framed record, parse it, validate its XF, and resolve its shared string
if it is a `LabelSst` — **re-reading, not caching the decoded value**, because
change 0300 deliberately retains SST offsets rather than decoded strings and this
design does not reopen that.

### Which refusals move: none, under one rule

The rule is ODS 0193/0195/0198's: *only successful scans are retained; no error
is ever cached*. It gives, without hand-waving:

- A worksheet carrying a defect never gets an index, because the scan that would
  have built it returned an error before EOF. Every query on that worksheet
  re-scans and refuses at the same record with the same message. This covers the
  three flagship worksheets whose shared-formula metadata this reader declines.
- A worksheet with no defect has nothing for a later query to refuse: the first
  scan validated every record to EOF, XF indices included.
- `max_worksheet_scan_bytes` and `max_worksheet_scan_records` are **per-query
  ceilings, not cumulative budgets** — verified in the tree: the XLS worksheet
  scan never charges `ExecutionContext`'s `Budget`, and `check_execution` tests
  cancellation only. A caller whose ceiling refuses the first scan gets that
  refusal on every query, because there is never an index. A query served from
  the index charges nothing it could exceed. This is the trap change 0592
  documented for DOCX and it does not bite here, but it had to be checked rather
  than assumed.

### What does change, and must be measured rather than assumed

**Source observations.** A second query on `54016.xls` today takes 19 of them
across its window fills; served from the index it would take the leading and
trailing fence, the single record's fill, and the three a shared-string resolve
takes — about six. The `SourceChanged` guarantee still holds, because
`ensure_current()` brackets the operation at both ends and `SourceVersion`
carries a revision counter, so an A-B-A mutation is still detected. What narrows
is the *window* in which a mid-operation mutation surfaces. Changes
[0279](changes/0279-cfb-operation-freshness-session-rejected.md) and
[0358](changes/0358-xls-worksheet-span-batching-rejected.md) were both rejected
in this neighbourhood, and both died on p95/p99 stability rather than on
semantics. This is the axis an implementation must measure first.

**Peak retained bytes.** Measured, not modelled: the walk added in part (1)
reports **38,950 cells** on `54016.xls` worksheet 0. At 16 bytes per slot that is
**623 KB**; at 24, **935 KB** — for **one** worksheet of a **984 KB** file. The
flagship has 16 worksheets. Retaining that for every worksheet a caller touches,
on an owner whose whole point is not to materialize the workbook, is the
bounded-memory contract's problem, and it is the survey's own falsification
criterion: *"falsified if … the retained weight breaks the bounded-memory
contract."*

### Predicted saving

A second query on the same worksheet costs a second complete scan today. Measured
exactly, as the difference between `second-cell` and `one-cell`:

| Fixture | extra reads | extra bytes | extra observations |
| --- | ---: | ---: | ---: |
| `54016.xls`, sheet 0 | 25 | 615,822 | 19 |
| flagship, sheet 1 | 7 | 37,907 | 10 |
| `WithCustomViews.xls`, sheet 1 | 1 | 414 | 4 |

Served from the index, a second query would issue one seek and one window fill
covering the record — the first fill of a scan is 512 bytes and a BIFF8 payload
is at most 8,224 — plus a shared-string resolve where the cell needs one.
**Modelled** at 2-3 reads and under 9 KB: on `54016.xls` about 10× fewer reads
and 70× fewer bytes, which clears the survey's *"at least 5× cheaper than a
re-scan"* bar on I/O. In instructions the saving is larger still: the second
query would parse one record instead of framing 37,929, against the 17,285,796 Ir
(70.5% of the query) change 0584 attributed to the scan.

### Admission gates a future implementation must clear

1. Byte-identical cell values for every position of every XLS fixture, first
   query and second, against a query on a fresh snapshot.
2. `Display`-equal refusals fixture by fixture, the three flagship worksheets
   included.
3. Reads, bytes and observations reported for both, with the observation
   narrowing stated in the record, not folded into a mean.
4. A weighted, bounded, evictable cache whose byte budget is a field of
   `SourceBackedLimits`, with a test that evicts under pressure and proves the
   evicted sheet's next query is value-identical, and with active walks pinning
   their index.
5. Paired p95/p99 timing that does not repeat 0279's and 0358's failure.

### Why it is not implemented

The brief's condition was *value-identical, and the retention accounting in
place*. The first half is achievable: the rule above holds every value and every
refusal where it is. The second half is not.

- **The accounting does not exist.** ADR 0005 requires semantic payloads to load
  into *"thread-safe weighted caches"* whose *"clean parsed values are
  evictable"*. A sweep of the record set finds no weighted, evictable cache
  anywhere in this repository: 0193, 0195, 0198 and 0592 are all non-evictable
  `OnceLock`s over data whose size is O(document structure), and 0005's XLSX
  row-start index is eager and O(cells already retained). A per-sheet index of
  0.6-1.0 MB over cells that are *not* retained is the first thing in this tree
  that would need the machinery, and building that machinery is a subsystem, not
  the smallest coherent change this brief authorized.
- **Part (1) removed the motivating case.** The scenario that made the index look
  large — reading many cells of one worksheet — is now one scan with zero
  retention, measured above at 32.3 bytes per cell against 617,079.3. What the index
  would still serve is *scattered, non-batchable, repeated* selected-cell queries
  on one snapshot. No selector in this repository measures that, no corpus
  evidence of such use exists, and the second-cell rows above are the only
  measurement of it that has ever been taken.

So the design is frozen here with its predicted saving on the record, and
`litchi-xls` gains no cache in this change.

## Limitations

**What is not claimed.** No speedup of any existing operation. `open`, `list`,
`one-cell` and full text are controls, and their instruction counts moved by at
most +0.146%, which is why the record calls them unchanged rather than improved.
No cold-cache, physical-device, remote or range-source, peak-RSS, allocation,
concurrency-scaling, real-producer or cross-platform result is claimed. No
claim-registry entry is made.

**The all-cells comparison is not a before-against-after leg.** It cannot be: the
before leg has no whole-sheet walk. Its two legs are the walk and a per-cell
reading using only pre-existing public API, in one binary and one build, which is
tighter pairing than a cross-leg comparison but measures a *new capability*
against *what a caller had to write*, not an optimization of an existing path.

**The per-cell extrapolation to a whole sheet is modelled.** 24.6 s for
`54016.xls` rests on three measured points (8, 64 and 256 positions) that are
linear to within 5%, not on a measured 38,950-position run.

**Two wall-clock review triggers fired** and are reported above with the
same-cell B/B floor beside them. They are on `open` and `list`, which this change
does not touch.

**The heavy scenarios' wall clock is below its own noise.** A same-binary floor
of 46.51% on the `54016` walk means no latency statement about the walk,
per-cell reading or full text is admissible from this capture beyond the ratios,
which exceed it by an order of magnitude.

**`visit_cells` reports what the worksheet stores, not a rectangle.** A position
with no record is not reported; a position stored twice is reported twice, in
stream order. A caller that wants the `DIMENSIONS` rectangle, blanks included,
must build it — the text projection still does that itself, and this change did
not unify the two.

**`scan_text_sheet` was deliberately not folded into the shared loop.** It has
its own semantics — it interprets `DIMENSIONS`, it refuses a `STRING` with no
pending `FORMULA` where `query_cell` skips it, and it accumulates into a
`HashMap` under the text budget — and full text already scanned each sheet once,
so there was nothing to win and a real risk of moving a refusal. Three scan loops
now exist in `source.rs` where there were two; that is a cost this change accepts
and names.

**A column above 255 is reachable by the walk and not by a query.**
`query_cell` returns `Ok(None)` for any column above `u8::MAX`, so a malformed
worksheet storing a record at column 256 or beyond would be reported by
`visit_cells` and unreachable through `cell()`. No fixture in the corpus does
this — the differential test asserts the walk's widest column is at most 255 on
every fixture it covers — but the asymmetry is real and is named here rather
than hidden.

**The retained sheet index is designed and not built.** Its predicted saving is
modelled from measured parts; nothing in part (2) has been executed.

**`54016.xls`'s 38,950 cells exceed the 29,608 cell parses** change 0584 counted
for the same worksheet, because `MulRk` and `MulBlank` records expand to several
cells each and 0584 counted records parsed, not cells produced. Both numbers are
correct for what they measure.

## Retained evidence

[`results/change-0605/README.md`](results/change-0605/README.md).
