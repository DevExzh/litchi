# Log sections for change 0595

Four paragraphs for the coordinator to merge, one per log, in the style of each
file's newest section. This change does not edit `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` or `ADR_COMPLIANCE.md` itself.

---

## For `HOTSPOTS.md`

## 0595 — retained lean XLS frame loop and cheaper eager SST walk

Retains change 0587 items XLS-1 (rank 9, change 0584's never-landed candidate 2)
and XLS-2b (rank 17). Five `ok_or` sites that built and dropped a 48-byte
`SourceBackedError::ResourceLimit` on every worksheet frame, every global record
and every extracted text cell became `let ... else`; `WorksheetScan::ensure`
split into an inlined resident test and an outlined `fill`; `check_execution`
inlined its uncancellable case; `frame_header` takes one bounds check; the
`SstCursor` maintains its logical position instead of re-summing every segment
behind it; fixed-width shared-string header fields are read by direct indexing
instead of a per-field `memcpy`; and the formatting-run walk is parameterised on
a sink so the source-backed walk stops allocating a `Vec` per string it
discards. One `54016.xls` open falls 253,317 → 193,739 ns (−23.5%), 7,240,329 →
5,375,911 Ir (−25.8%) and 1,165,313 → 869,124 cycles (−25.4%); one-cell falls
1,011,306 → 832,605 ns (−17.7%) and 24,676,105 → 18,602,879 Ir (−24.6%);
`WithCustomViews.xls` falls 6.2–13.0% at p50 on every operation and mode.
`drop_in_place<SourceBackedError>` and `SstCursor::read_exact` are exactly zero
on all three fixtures after the change. The flagship moves only 2–5% in counts
and its wall clock sits inside the measured ±3.1% floor, so no speedup is
claimed for it. Framing overhead on the `54016` one-cell query falls from 46.3%
of the operation to 38.5% and remains the largest single term; XLS-2 (lazy SST
indexing) and XLS-3 (retained sheet index, whole-sheet iterator) still require
their own frozen design records. OLE2/OOXML optimization remains active; ODF is
deferred until completion and iWork excluded. [Change and
limitations](0595-xls-frame-loop-and-sst-walk.md); [retained
evidence](results/change-0595/README.md).

---

## For `GOAL_AUDIT.md`

## 0595 — retained lean XLS frame loop and cheaper eager SST walk

Closes two rows of change 0587's ranked queue by GOAL step 1 (eliminate
unnecessary work) and step 3 (unnecessary allocation and copying), with no step
skipped: no layout change, no algorithm substitution beyond replacing an O(n)
re-summation with a maintained running total, no parallelism and no SIMD. The
decision rules are satisfied in order — before measurements captured on a
read-only detached checkout of the base, hypothesis and mechanism stated per
site, smallest coherent change, correctness/preservation/adversarial evidence,
after measurements with identical setup. Evidence tiers: **measured** for the 18
logical-counter cells, the six callgrind isolation pairs, the six `perf stat`
pairs and the 72 timing rounds; **modelled** for the two `SourceTextSheet::insert`
sites, which sit on the text-extraction path that `xls_source_attribution`
still has no selector for — change 0587's recorded measurement blocker, now
blocking a second batch. The audit row "apply layout, cache, or SIMD tuning only
from measured hot loops" is untouched. The standing instruction to price
pointer-chase work in cycles is honoured: native cycles fall 25.42% and 18.44%
on the `54016` open and one-cell against callgrind's 25.75% and 24.61%, and IPC
falls on five of the six cells because what was removed was the most
superscalar-friendly work in the loop. Host quiescence is not established — the
load average was 44.41 to 40.45 on 32 cores — so the noise floor is measured in
the same window (−2.64% to +3.06% at p50 over 36 same-binary comparisons)
rather than assumed.
OLE2/OOXML optimization remains active; ODF is
deferred until completion and iWork excluded. [Change and
limitations](0595-xls-frame-loop-and-sst-walk.md); [retained
evidence](results/change-0595/README.md).

---

## For `REPORT.md`

## 0595 — retained lean XLS frame loop and cheaper eager SST walk

Two files changed in `litchi-xls`, seven edits, `performance_claim: none`.
Paired A1 B1 B2 A2 medians over 500 samples per round: `54016.xls` open −23.79%
/ −23.24%, list −22.13% / −21.29%, one-cell −18.08% / −17.26% (`owned-readat`),
and −19.92% / −20.03%, −23.57% / −23.06%, −17.79% / −17.36% (`file-source`), all
six cells agreeing between directions to within 0.9 percentage points;
`WithCustomViews.xls` between −6.23% and −13.02%, its `file-source` one-cell
being the noisiest row in the table and carrying both extremes of the floor;
`ConditionalFormattingSamples.xls` between −0.59% and −3.90%, which is inside
the measured ±3.1% floor and is reported as not separable. Tails move with the
medians (`54016` owned one-cell p99 1,029,966 → 840,355 ns on adjacent rounds);
two host artifacts, both on the *before* leg in this window, are named rather
than smoothed. Two regressions are reported rather than netted
into a mean: an inlining artifact
outside either mechanism (`CellAlignment::parse` outlined from `parse_xf`) plus
`MeasuredText::consume`, together about 1.0% of the `54016` open. Controls: all
18 logical-counter cells and all 18 semantic projections identical between the
legs, the three open rows reproducing changes 0565, 0574 and 0576 exactly, and
the before instruction profile reproducing change 0584 to within 0.045% on the
opens. Gates: `cargo fmt --all --check`, `cargo clippy -p litchi-xls
--all-targets`, `cargo test -p litchi-xls` (1,382 passed, 0 failed, 1
pre-existing ignored doctest) and `cargo doc -p litchi-xls --no-deps`, all
clean. No cold-cache, physical-device, range-source, RSS, allocation-profile,
concurrency-scaling, real-producer or cross-platform result is claimed.
OLE2/OOXML optimization remains active; ODF is deferred until completion and
iWork excluded. [Change and limitations](0595-xls-frame-loop-and-sst-walk.md);
[retained evidence](results/change-0595/README.md).

---

## For `ADR_COMPLIANCE.md`

## 0595 — retained lean XLS frame loop and cheaper eager SST walk

ADR 0005 is the governing record and is satisfied unchanged: every
`ResourceLimit` still identifies its resource, observed value and ceiling, and
`max_worksheet_scan_records`, `max_worksheet_scan_bytes`, `max_global_records`,
`max_text_bytes` and `max_text_cells` are compared against the same values at
the same points — only the moment at which the error *value* is constructed
moved, from unconditionally to on the path that returns it. ADR 0006's
preservation clause is not engaged (nothing here is on a write path) and ADR
0003's transaction boundary is not engaged. No new `unsafe`, no weakened
malformed-input defence, no hidden global Rayon pool, no ambient I/O, no public
leakage of archive types, raw locks or executors; no public API changed. Source
freshness fences are unchanged and counted: `version()` observations are
identical between the legs in all 18 cells. **One behavioural difference is
recorded rather than left to be discovered:** the source-backed shared-string
walk no longer attempts `try_reserve_exact` for formatting runs, so under
allocator exhaustion it no longer produces that typed allocation refusal, and on
a string whose runs are also malformed the run check now wins where the
allocation message used to. The reservation guarded a `Vec` the walk discarded —
at most 65,535 runs × 4 bytes — nothing newly unbounded is retained, and the
eager `SharedStringTable`, which keeps its runs, still takes the reservation and
still refuses. It is a bounded allocation that no longer happens, not a ceiling
that was relaxed, and no input decides it. No
input-dependent refusal moves: the SST index is byte-identical over the 121
fixtures that carry an SST
against a digest captured on the untouched base commit, and change 0576's corpus
differential still compares 17,434 entries and four identical refusals with
identical messages. OLE2/OOXML optimization remains active; ODF is deferred
until completion and iWork excluded. [Change and
limitations](0595-xls-frame-loop-and-sst-walk.md); [retained
evidence](results/change-0595/README.md).
