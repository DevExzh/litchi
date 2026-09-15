# Log sections for change 0605

Four paragraphs for the coordinator to merge, one per log, in the style of each
file's newest section. This change does not edit `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` itself.

---

## For `HOTSPOTS.md`

## 0605 — retained whole-sheet XLS walk; retained sheet index frozen as a design

Takes change 0587 item XLS-3 (rank 18), whose two halves the survey separated,
and lands the first: `SourceBackedWorksheet::visit_cells` reports every stored
cell of a worksheet from **one** validated scan, where reading a worksheet cell
by cell cost one complete scan of its substream per cell. `query_cell`'s frame
loop became `scan_worksheet`, generic over a private `CellSink`, with
`process_cell` moved behind it unchanged and a second sink that reports every
cell; the only branch the two sinks take differently is the one `query_cell`
already had, holding a string-valued `FORMULA` back for its `STRING` result. On
`54016.xls` the walk reports **38,950 cells for 16,145 reads and 1,256,139
bytes** against **157,972,311 bytes for 256 cells** read one query at a time —
32.3 bytes per cell against 617,079.3 (19,134×) and 9,146 instructions per cell
against 13,330,564 (1,457×); the crossover is about three dozen cells on every
fixture measured. Nothing else moved: reads, bytes, source observations, `len`
calls and seeks are identical to the byte on `open`, `list` and `one-cell` over
three fixtures and both source modes, and their instruction counts moved by at
most +0.146%. This closes change 0568's third limitation, carried unaddressed
for 37 records, and unblocks the measurement gap change 0587 recorded:
`xls_source_attribution` now has `all-cells`, `full-text` and `second-cell`
operations. The **second half — the snapshot-scoped retained sheet index — is
frozen as a design and not implemented**: its measured weight is 623 KB–935 KB
for one worksheet of a 984 KB fixture, ADR 0005's weighted, bounded, evictable
cache has no implementation anywhere in this repository, and part (1) removed the
index's motivating case. XLS-2 (lazy SST indexing) still requires its own frozen
design record; XLS-4, 8 and 10 are now cheaper to size because the walk exists.
OLE2/OOXML optimization remains active; ODF is deferred until completion and
iWork excluded. [Change and
limitations](0605-xls-retained-sheet-index.md); [retained
evidence](results/change-0605/README.md).

---

## For `GOAL_AUDIT.md`

## 0605 — retained whole-sheet XLS walk; retained sheet index frozen as a design

GOAL step 1 (eliminate unnecessary work) in its purest form: the work removed is
N−1 complete validated scans of a worksheet substream, and nothing about the
remaining scan changed. No step was skipped — no layout change, no algorithm
substitution, no parallelism, no SIMD — and no validation was moved, relaxed or
deferred, because the walk runs to the worksheet's EOF exactly as a selected-cell
query does; an early exit would be change 0574's opportunity 6, which ADR 0005's
mandatory-validation clause rejects. The decision rules are satisfied in order:
before measurements captured from the shared read-only checkout of the base,
hypothesis and mechanism stated, smallest coherent change, correctness and
adversarial evidence, after measurements with identical setup. Evidence tiers:
**measured** for the 42 logical-counter cells, the 24 walk-against-per-cell
cells, the 18 callgrind isolation pairs (36 annotated profiles), the 18
`perf stat` pairs, the 120 timing runs across four rounds and the corpus
differential over every XLS fixture; **modelled** for the 24.6 s whole-sheet
per-cell extrapolation, which rests on three measured points that are linear to
within 5%, and for every number in the part (2) design. The rule that a measurement blocker is itself work
was applied rather than noted: change 0587 recorded that no source-backed XLS
all-cells or full-text selector existed, so this change built them, and their
all-cells oracle is a differential that fails the run on any disagreement between
the walk and a selected-cell query. The standing instruction to price
pointer-chase work in cycles is honoured, and it is what caught the one real
problem: the first shared loop cost +1.37% instructions and +1.36% cycles on the
`54016` one-cell query, `#[inline]` hints and a restored branch shape did not
recover it, and folding the sink's four constant arguments into one `ScanContext`
did, to +0.011%; the three trial profiles are retained rather than discarded.
Host quiescence is not established — eight measurement agents shared 32 cores,
1-minute load 8.27 — so the floor is measured in the same window: A/A p50 0.71%,
max 2.14% over the 18 cells that exist on both legs, but B/B up to 46.51% on the
heavy new scenarios, which is why no wall-clock statement is made about them
beyond their ratios. OLE2/OOXML optimization remains active; ODF is deferred
until completion and iWork excluded. [Change and
limitations](0605-xls-retained-sheet-index.md); [retained
evidence](results/change-0605/README.md).

---

## For `REPORT.md`

## 0605 — retained whole-sheet XLS walk; retained sheet index frozen as a design

Three files changed in `litchi-xls` and the standalone attribution harness,
`performance_claim: none`. The headline is a new capability, not a speedup of an
existing path, so it is reported as a ratio inside one binary and one build:
`visit_cells` against one `cell_value_by_index` per position, the latter using
only API that predates this change. `54016.xls` worksheet 0 — 38,950 cells,
16,145 reads, 1,256,139 bytes, p50 20.71 ms — against 256 positions read one at a
time — 6,529 reads, 157,972,311 bytes, p50 161.99 ms; `WithCustomViews.xls`
sheet 0, 3,325 cells for 257,066 bytes against 10,499,248 for 256; the flagship's
sheet 11, 347 cells for 600,726 bytes against 9,558,820 for 256. Controls: all 18
before-against-after counter cells identical in reads, bytes, observations, `len`
calls and seeks, with identical semantic projections, and instruction counts
+0.011% to +0.146% across the six isolation pairs. Full text is a control too and
is unchanged — it already scanned each sheet once — including the flagship's
refusal, which the harness now records as an outcome rather than a failure and
which is byte-identical across legs. **Two of 36 paired wall-clock comparisons
exceed the +5% review trigger** and are reported rather than netted into a mean:
`WithCustomViews.xls` `file-source` open (+7.19%) and list (+6.75%), both in one
direction only, both inside the same cell's 6.67% and 6.14% B/B floor, both on
operations that never enter the worksheet frame loop this change touched, and
both on the smallest absolute times in the matrix (34.7 µs, where 2 µs is 6%).
The heavy new scenarios carry a same-binary floor up to 46.51% at 30 samples, so
their evidence is the deterministic counters and instruction counts, not their
wall clock. Gates: `cargo fmt --all --check`, `cargo clippy` on both touched
crates, `cargo test -p litchi-xls` (72 binaries, 1,390 passed, 0 failed, 1
pre-existing ignored doctest), the harness's own 11 tests, and `cargo doc -p
litchi-xls --no-deps`, all clean. Part (2), the snapshot-scoped retained sheet
index, is designed and not built. No cold-cache, physical-device, range-source,
RSS, allocation-profile, concurrency-scaling, real-producer or cross-platform
result is claimed. OLE2/OOXML optimization remains active; ODF is deferred until
completion and iWork excluded. [Change and
limitations](0605-xls-retained-sheet-index.md); [retained
evidence](results/change-0605/README.md).

---

## For `ADR_COMPLIANCE.md`

## 0605 — retained whole-sheet XLS walk; retained sheet index frozen as a design

ADR 0005 is the governing record and part (1) is the rare case that leaves it
entirely alone: the walk adds no cache, retains nothing between calls, loads
nothing eagerly, and performs the same lazy scan `query_cell` performed, exposed
once instead of once per cell. Every check is taken in the same order against the
same bytes; `max_worksheet_scan_records` and `max_worksheet_scan_bytes` are
compared at the same points; the leading cancellation check and `ensure_current`
fence and the trailing pair are the ones `query_cell` and `finish_query` took,
and `version()` observations are identical between the legs in all 18 paired
cells. ADR 0003 is untouched — `SourceBackedWorksheet` is still a two-word
lifetime-free handle — and ADR 0006 is not engaged, since nothing here is on a
write path. No new `unsafe`, no weakened malformed-input defence, no hidden global
Rayon pool, no ambient I/O, no public leakage of archive types, raw locks or
executors; `CellSink`, `ScanContext`, `TargetCell` and `VisitCells` are private
and the only public surface added is two methods over already-public types. The
walk cannot be used to skip validation: it always runs to the worksheet's EOF and
the visitor cannot stop it. **Two behavioural notes are recorded rather than left
to be discovered.** First, a `try_reserve` was added before the `Vec::push` of a
`FORMULA` `STRING` continuation, which the text path already had and `query_cell`
did not; it can only convert an allocator abort into
`SourceBackedError::Allocation`, and the counters show it is unreachable on every
fixture measured. Second, `query_cell` returns `Ok(None)` for any column above
`u8::MAX`, so a malformed worksheet storing a record at column 256 or beyond
would be reported by `visit_cells` and unreachable through `cell()`; no corpus
fixture does this and the differential test asserts the walk's widest column is
at most 255 on every fixture it covers. Refusal identity is proved by `Display`
comparison, not by variant matching, on one synthetic defect and on the three
real flagship worksheets whose shared-formula metadata this reader declines.
**Part (2) is where ADR 0005 bites and is the reason nothing was built**: its
lazy-payload clause permits a first-use index, but its "thread-safe weighted
caches" whose "clean parsed values are evictable" has no implementation anywhere
in this repository — changes 0193, 0195, 0198 and 0592 are all non-evictable
`OnceLock`s over data whose size is O(document structure), and change 0005's XLSX
row-start index is eager over cells already retained — while the measured weight
here is 623 KB–935 KB for one worksheet of a 984 KB fixture. The design records
the rule that would keep every refusal in place (retain only complete successful
scans), the one thing that would measurably change (the freshness-observation
window, the axis on which changes 0279 and 0358 were rejected), and the five
admission gates a future implementation must clear. OLE2/OOXML optimization
remains active; ODF is deferred until completion and iWork excluded. [Change and
limitations](0605-xls-retained-sheet-index.md); [retained
evidence](results/change-0605/README.md).
