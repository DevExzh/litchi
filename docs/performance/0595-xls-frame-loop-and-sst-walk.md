# 0595: a leaner XLS frame loop and a cheaper eager SST walk

Status: retained. `performance_claim: none` — this record carries deterministic
counts, callgrind isolation pairs, native hardware counters and paired medians
in both directions with a measured noise floor, not a claim-registry entry.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements items **XLS-1** (rank 9) and **XLS-2b** (rank 17) of change
[0587](0587-remaining-opportunity-survey.md), both of which that survey marked
"low risk, no design record needed". XLS-1 is change 0584's candidate 2, which
was proposed and never landed. Neither item moves a refusal, changes an output
byte, weakens a fence or relocates a typed limit; the one behavioural difference
the work uncovered is named under *Validation preserved* rather than left for a
reader to find.

## What was changed

Two files, `crates/litchi-xls/src/workbook/source.rs` and
`crates/litchi-xls/src/records.rs`. Seven edits, in two groups.

### XLS-1: the worksheet frame loop stops building errors it does not return

**Five eager error values became `let ... else`.** `Option::ok_or` takes its
error **by value**, so every one of these sites built a 48-byte
`SourceBackedError::ResourceLimit` and — on the overwhelmingly common `Some`
path — dropped it again. `SourceBackedError` carries `String` variants, so the
drop is a real `drop_in_place` call, not a no-op. Change 0584 counted 75,858
such drops in one `54016.xls` one-cell query: exactly 2 × 37,929 frames. The
sites are `WorksheetScan::next_frame` (twice, the scanned-record and
scanned-byte counters), `parse_globals` (the global-record counter) and
`SourceTextSheet::insert` (twice, the retained-text-byte and text-cell
counters).

The obvious spelling, `ok_or_else`, does not survive the gates: clippy's
`unnecessary_lazy_evaluations` rejects a closure whose body is a struct literal
of simple fields, and this workspace denies warnings. That is almost certainly
why 0584's candidate never landed. `let Some(x) = ... else { return Err(...) }`
expresses the same thing, passes clippy, and reads better at the point of use —
the limit value now sits next to the `return` that produces it.

**`ensure` was split into an inlined resident test and an outlined fill.**
`WorksheetScan::ensure(need_end)` returns immediately when the bytes are already
in the window, which is what happens on every frame and every payload of a sheet
whose records sit inside the current fill. It was a call returning its `Result`
through memory — change 0587 measured `SourceBackedError` at 48 bytes and
`Result<WorksheetFrame, SourceBackedError>` at 48. It is now `#[inline]` and
carries the resident test only; the body that drains, reserves, zero-fills and
reads is `#[inline(never)] fn fill`. Nothing about the order or the identity of
what `fill` does changed; the one textual difference is that `fill` recomputes
`filled_end()`, on the cold path, instead of receiving it.

**`check_execution` inlines its uncancellable case.** It runs twice per framed
record — once in `next_frame`, once in `consume_payload` or `skip_payload` — and
most callers pass no `ExecutionContext`. `self.execution.map_or(Ok(()), ...)`
became a `match` with `#[inline]`, so the null case is a branch rather than a
call returning a `Result` through memory.

**`frame_header` takes one bounds check instead of four.** `at` is already
clamped to `self.window.len()`, so `&self.window[at..at + 4]` cannot overflow
and panics on exactly the residency failure the four element indexes panicked
on — a state `ensure` makes unreachable.

### XLS-2b: the eager SST walk stops paying per string

**`SstCursor` maintains its logical position instead of recomputing it.**
`logical_position` summed the lengths of every segment behind the cursor, and
the SST scan calls it **twice per shared string** to record an entry's `start`
and `end`. On `54016.xls` — 7,893 strings over 28 segments — that is the
dominant cost of `scan_shared_string_records` itself. A `logical_base` field,
advanced by `advance_segment` and restored by the new `seek`, makes it O(1). The
old summed expression survives as a `debug_assert_eq!` inside
`logical_position`, so every debug test run recomputes it and compares.

**Fixed-width header fields are read by direct indexing.** `SstCursor::read_exact`
issued a `copy_from_slice` — a `memcpy` call — per two- or three-byte header
field, 7,900 times in one `54016.xls` open. `take_resident::<N>()` returns the
`N` bytes when they lie inside the current segment, which is the case
`read_exact`'s loop handled with one copy and no `advance_segment`, and
`read_u16_continued` and `read_u32_continued` fall through to the unchanged
`read_exact` otherwise. `read_exact` remains the only implementation of
continuation crossing.

**The formatting-run walk is parameterised on a sink.** `read_formatting_runs`
allocated a `Vec<SharedStringFormatRun>` per shared string;
`walk_one_shared_string` — the only caller on the source-backed path, in **both**
of its instantiations — discarded it. Following change 0576's pattern, the
validation now lives once in `walk_formatting_runs<S: FormatRunSink>`. The eager
`SharedStringTable::parse_segments`, which genuinely keeps the runs,
instantiates it with `Vec` through the unchanged `read_formatting_runs` alias;
the source-backed walk instantiates it with `MeasuredRuns`, which retains
nothing.

## Why it is sound

**Every check is the same check, in the same order, with the same message.** The
`let ... else` rewrites move *when the error value is constructed*, never when
the branch is taken: `checked_add` is evaluated exactly where it was, the
comparison against the limit is the next statement, and the `resource`,
`observed` and `maximum` fields are byte-for-byte the ones the old expression
built. ADR 0005 requires that "limit errors identify the resource, observed
value, limit, and object path"; all four sites still do.

**The function splits are splits.** `ensure`/`fill`, `check_execution`'s `match`
and `read_formatting_runs`/`walk_formatting_runs` are refactors with identical
control flow; `#[inline]` and `#[inline(never)]` are hints to the optimizer with
no semantic content. No `unsafe` was added; there is none in either file.

**`take_resident` is a case split, not a second parser.** It returns `Some` only
when `offset + N` is inside the current segment — precisely the iteration in
which `read_exact`'s loop performs one `copy_from_slice` of `N` bytes, advances
`offset` by `N`, and calls neither `advance_segment` nor `current()` a second
time. Every other input, including every `Continue`-crossing header field, still
goes through `read_exact` unchanged, so the refusals that framing owns
(`UnexpectedEndOfStream` with its per-field context string) are produced by the
same code as before.

**`logical_base` is proved, not asserted.** `advance_segment` adds the length of
the segment it leaves (past the last segment `current()` is empty and adds
nothing, which is what `take(segment_index)` did when it ran past the end);
`seek` restores it with the index and offset, which is the only rewind in the
walk — `measure_characters`'s cold path. The `debug_assert_eq!` recomputes the
summed form on **every** call in a debug build, so the 1,382 tests in this crate,
including the corpus differential's 17,434 entries over 121 fixtures, are each a
proof that the maintained base equals the sum.

**Contracts untouched.** No refusal moves in the malformed-input sense: the
pinned corpus index and change 0576's differential both show the same entries and
the same four refusals with the same messages. No output byte changes: nothing
here is on a write path. No fence moves: the source-version observations are
unchanged and counted (see *Measured*). No limit is relaxed:
`max_worksheet_scan_records`, `max_worksheet_scan_bytes`, `max_global_records`,
`max_text_bytes` and `max_text_cells` are compared against the same values at
the same points. ADR 0006's preservation clause is not engaged; ADR 0003's
transaction boundary is not engaged.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Both legs are `tools/perf-baseline`'s `xls_source_attribution`
built `--release --locked`; the before leg from a shared read-only detached
checkout of the base `08d968f8e` with its own `CARGO_TARGET_DIR`, the after leg
from this branch's worktree. Binary hashes, fixture hashes and the quiescence
log are in [`environment.json`](results/change-0595/environment.json). **Host
quiescence is not established**: eight measurement agents shared this machine
and the load average was 44.41 when the counter window opened and 40.45 when the
timing window closed, on 32 cores. That is why the floor below is measured
rather than assumed. Every measured child ran pinned to CPU 15 with ASLR
disabled.

### The control: not one byte of I/O moved

Eighteen cells — three fixtures × two in-memory source modes × `open`, `list`,
`one-cell` — and every logical counter is identical between the legs, as is the
harness's implementation projection of the selected cell and the worksheet
names:

| fixture | reads | read bytes | `version()` | `len()` | seeks |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` open | 53 | 565,201 | 29 | 1 | 0 |
| `WithCustomViews.xls` open | 16 | 110,242 | 22 | 1 | 0 |
| `54016.xls` open | 40 | 317,171 | 25 | 1 | 0 |
| `54016.xls` one-cell | 65 | 932,993 | 43 | 1 | 0 |

The three open rows reproduce change 0576's retained figures exactly, which
themselves reproduce 0565's and 0574's. This change removes CPU work only.

### Instructions per operation, callgrind isolation pairs

A large-sample and a small-sample child are differenced and divided by the extra
operations, so everything that runs once per child — the harness's SHA-256 of
the input, its eager oracle projection — cancels.

| fixture | operation | before | after | change |
| --- | --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | open | 2,378,029 | **2,327,817** | **−2.11%** |
| `ConditionalFormattingSamples.xls` | one-cell | 2,625,686 | **2,550,722** | **−2.86%** |
| `WithCustomViews.xls` | open | 915,766 | **807,974** | **−11.77%** |
| `WithCustomViews.xls` | one-cell | 937,854 | **827,140** | **−11.81%** |
| `54016.xls` | open | 7,240,329 | **5,375,911** | **−25.75%** |
| `54016.xls` | one-cell | 24,676,105 | **18,602,879** | **−24.61%** |

The before column reproduces change 0584's independently captured profile to
0.045%, 0.026% and 0.011% on the three opens, and to 0.34%, 0.04% and 0.61% on
the three one-cell queries. That agreement is the strongest available evidence
that this record and 0584 measure the same thing.

### Where the instructions went, self Ir per operation

`54016.xls`, the fixture both items were sized on. Every row is a symbol whose
self cost moved by more than 0.2% of the before total.

| symbol | one-cell before | after |
| --- | ---: | ---: |
| `WorksheetScan::ensure` | 1,758,790 | **0** (inlined; `fill` is 2,754) |
| `WorksheetScan::next_frame` | 4,248,048 | **2,806,836** |
| `drop_in_place<SourceBackedError>` | 1,030,627 | **0** |
| `records::scan_shared_string_records` | 1,164,269 | **238,178** |
| `SstCursor::read_exact` | 624,100 | **0** |
| `SstCursor::read_formatting_runs` | 236,890 | **0** (`walk_formatting_runs` is 165,818) |
| `records::walk_one_shared_string` | 836,795 | **742,217** |
| `WorksheetScan::read_payload` | 1,488,468 | **1,456,170** |
| `__memcpy_avx_unaligned_erms` | 1,913,201 | **1,771,026** |
| `MeasuredText::consume` | 719,535 | 759,098 |
| `number_format::codec::parse_xf` | 297,066 | 237,737 |
| `alignment::CellAlignment::parse` | 0 | 95,480 |

Four of these deserve naming rather than rounding away.

`drop_in_place<SourceBackedError>` is **exactly zero** on all three fixtures and
both operations. That is the precise claim of XLS-1's first half: the only
`SourceBackedError` values a successful source-backed read used to build were
the ones it threw away, and none are built now.

`ensure` is zero because it is inlined, not because it is gone; the reads it
used to issue are in `fill`, and all of them together are 2,754 Ir — the query
issues 25 more positional reads than the open, per the counter table above.
`next_frame` absorbed the inlined resident test and still fell
by 1,441,212 Ir, and `read_payload` absorbed it and fell by 32,298.

`read_exact` is **exactly zero** on all three fixtures: every shared-string
header field in this corpus lies inside one segment, so the fallback never runs
on it. The 142,174 Ir that left `memcpy` is the same removal seen from the other
side — callgrind charges `rep movsb` per byte, so that figure is an upper bound.

`parse_xf` and `CellAlignment::parse` are an **inliner artifact and a small
regression**: `CellAlignment::parse` used to be inlined into `parse_xf` and is
now outlined, and the pair costs 36,151 Ir more than `parse_xf` alone did. It is
in neither item's mechanism; perturbing the module changed an inlining decision
elsewhere. `MeasuredText::consume` likewise costs 39,563 Ir more. Together they
give back 1.0% of the 54016 open, inside a 25.75% saving.

### Cycles and IPC, native `perf stat`, same isolation method

Change 0579 established that instruction share mis-ranks pointer-chase work, so
the change is also priced natively. Callgrind and `perf` disagree in level
because callgrind counts `rep movsb` per byte; they agree in direction
everywhere.

| fixture | operation | cycles before | cycles after | change | instructions | IPC |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `Conditional…` | open | 353,412 | **347,574** | **−1.65%** | −4.22% | 3.48 → 3.39 |
| `Conditional…` | one-cell | 417,411 | **388,127** | **−7.02%** | −5.40% | 3.39 → 3.45 |
| `WithCustomViews` | open | 150,383 | **129,779** | **−13.70%** | −17.77% | 4.14 → 3.95 |
| `WithCustomViews` | one-cell | 155,652 | **133,185** | **−14.43%** | −17.49% | 4.14 → 3.99 |
| `54016` | open | 1,165,313 | **869,124** | **−25.42%** | −31.91% | 5.08 → 4.64 |
| `54016` | one-cell | 4,522,126 | **3,688,400** | **−18.44%** | −27.51% | 4.90 → 4.35 |

IPC falls on five of the six cells, and that is the expected shape rather than a
warning: what was removed — six-store error writes that are never read, a
summation over a segment array in L1, a two-byte `memcpy` call — is the most
superscalar-friendly work in the loop. Branch misses on the `54016` one-cell
fall from 4,589 to 2,276. The cycle column is the noisiest here because the host
was at load 44, and the two cells where it disagrees most with its own
instruction column (`Conditional…` open at −1.65% and one-cell at −7.02%) are
the two smallest operations measured; the instruction columns, which are
deterministic, agree with the callgrind table's direction on every cell.

### Wall clock, paired, A1 B1 B2 A2

`p50` of the measured operation in nanoseconds, 500 samples per round after 50
warmups, four rounds in the order before, after, after, before. `dir1` is the
first before→after pair and `dir2` the second.

| fixture | mode | operation | before p50 | after p50 | dir1 | dir2 |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| `54016` | owned | open | 253,317 | **193,739** | **−23.79%** | **−23.24%** |
| `54016` | owned | list | 253,821 | **198,706** | **−22.13%** | **−21.29%** |
| `54016` | owned | one-cell | 1,011,306 | **832,605** | **−18.08%** | **−17.26%** |
| `54016` | file | open | 265,194 | **212,219** | **−19.92%** | **−20.03%** |
| `54016` | file | list | 270,054 | **207,086** | **−23.57%** | **−23.06%** |
| `54016` | file | one-cell | 1,040,404 | **857,538** | **−17.79%** | **−17.36%** |
| `WithCustomViews` | owned | open | 31,410 | **27,585** | **−12.15%** | **−12.21%** |
| `WithCustomViews` | owned | list | 32,523 | **28,435** | **−13.02%** | **−12.11%** |
| `WithCustomViews` | owned | one-cell | 32,440 | **28,700** | **−12.69%** | **−10.35%** |
| `WithCustomViews` | file | open | 39,973 | **35,598** | **−11.94%** | **−9.96%** |
| `WithCustomViews` | file | list | 40,556 | **36,595** | **−9.43%** | **−10.10%** |
| `WithCustomViews` | file | one-cell | 41,322 | **37,662** | **−11.41%** | −6.23% |
| `Conditional…` | owned | open | 74,565 | 72,896 | −1.93% | −2.54% |
| `Conditional…` | owned | list | 75,308 | 73,616 | −2.74% | −1.76% |
| `Conditional…` | owned | one-cell | 85,452 | 83,968 | −1.51% | −1.97% |
| `Conditional…` | file | open | 93,228 | 90,248 | −2.49% | −3.90% |
| `Conditional…` | file | list | 95,041 | 92,040 | −3.44% | −2.87% |
| `Conditional…` | file | one-cell | 107,233 | 106,258 | −1.23% | −0.59% |

**The floor, measured in the same window.** Across all 36 same-binary
comparisons (a2 against a1, b2 against b1) the excursion at p50 is **−2.64% to
+3.06%**. Call the floor ±3.1%, on a host at load 44 with eight measurement
agents running, against this machine's standing figure of about 4% at p50. Both
extremes are the *same cell*, `WithCustomViews` `file-source` one-cell, which is
therefore the least trustworthy row in the table and the only one whose two
directions disagree by more than 2.4 percentage points.

**The flagship's rows are inside that floor and are not claimed as a speedup.**
`ConditionalFormattingSamples.xls` moves −0.59% to −3.90% against a ±3.1% floor.
Its instruction and cycle counts are deterministic and do fall (−4.22%
instructions on the `owned-readat` open), so the change is not *neutral* there;
the wall-clock evidence is simply not separable from noise, and it is reported as
such. The `54016` rows are 5.6 to 7.7 times the floor and agree between the two
directions to within **0.9 percentage points on all six cells**. The
`WithCustomViews` rows are 2.0 to 4.2 times the floor and agree to within 1.0
point on four of six, 2.3 on a fifth, and 5.2 on the noisy cell named above.

**Tails move with the medians.** On the `owned-readat` `54016` open, p95 falls
from 268,252 to 201,631 ns (a1 against b1, the adjacent rounds); on the
one-cell, p95 falls from 1,024,656 to 836,945 and p99 from 1,029,966 to 840,355.
**Two host artifacts are named rather than smoothed, and both land on the
*before* leg in this window**: the a1 round of the `54016` `owned-readat` open
has a p99 of 3,279,370 ns against a p95 of 268,252 — one sample — and the a1
round of the flagship `file-source` one-cell has a mean 20.1% above its median.
That is what a load-44 host does to a 100-microsecond operation; p50 and p95 are
stable across rounds on every cell, and the per-round table is retained so this
is visible.


## Correctness evidence

**A new corpus oracle, pinned across the change.** Change 0576's differential
proves the two instantiations of the SST walk agree with *each other*; it cannot
see a change that moves both, which is exactly what this change could have done.
`the_sst_index_over_the_corpus_is_pinned` digests every segment locator
(`source_offset`, `logical_offset`, `len`) and every entry boundary (`start`,
`end`) of every XLS fixture that carries an SST, per fixture and then over the
corpus, and asserts the total. **The constant was captured on the untouched base
commit, before any production edit** — the test was added to a worktree whose
two production files were still the base's (`git checkout` of both), run, and
only then was the change applied — and the test passes unchanged on this branch:
121 fixtures, digest `0x9cb14f5daa02eebc`. The test does not exist at
`08d968f8e`, so reproducing `sst-index-before.txt` means repeating that
sequence, which the packet `README.md` spells out. Both listings and their
`diff` are retained
([`sst-index-before.txt`](results/change-0595/sst-index-before.txt),
[`sst-index-after.txt`](results/change-0595/sst-index-after.txt),
[`sst-index-diff.txt`](results/change-0595/sst-index-diff.txt)); they are
identical line for line, including the four refusals and their messages.

**Change 0576's differential, unchanged and still green.**
`every_sst_fixture_indexes_identically_both_ways`: 126 fixtures, 121 with an
SST, 117 indexed identically by both instantiations over **17,434 entries**, 4
refused identically with identical messages. Those counts are exactly the ones
0576 retained.

**The `debug_assert_eq!` in `logical_position`** recomputes the old O(n) sum on
every call in every debug build, so all 1,382 tests — the corpus walk included —
check the maintained base rather than trusting it.

**One adversarial test was added**, because the corpus does not reach the case
that matters most here.
`a_formatting_run_split_across_a_continue_record_still_indexes_exactly` puts a
formatting run's `character_index` across a `Continue` boundary and a second
shared string after it. A formatting run is the *only* shared-string field that
can straddle a record boundary — `ensure_current` keeps the header, the
rich-text count and the extension length inside one record — so it is the only
place the scan can exercise `take_resident` returning `None`, `read_exact`'s
continuation loop, and `advance_segment` maintaining `logical_base` mid-string.
The second string proves the base is still right after the crossing. **The test
was run against the untouched base commit as well and produces the same
`ok [(8, 19), (19, 24)]`**, so it is a cross-version oracle rather than a
snapshot of new behaviour.

**Gates**, all run in this worktree, tails in
[`gates.txt`](results/change-0595/gates.txt):

- `cargo fmt --all --check` — clean.
- `cargo clippy -p litchi-xls --all-targets` — clean. Workspace lints are deny,
  and `unnecessary_lazy_evaluations` is the lint that shaped the `let ... else`
  spelling above.
- `cargo test -p litchi-xls` — 72 test binaries, **1,382 passed, 0 failed, 1
  ignored**. The ignored test is a pre-existing `#[ignore]` doctest in
  `writer/core/codec/worksheet.rs`, a file this change does not touch.
- `cargo doc -p litchi-xls --no-deps` — clean; rustdoc lints are deny.

**The harness's own oracle** agrees on all 18 counter cells: the worksheet count,
the worksheet names and the selected cell's typed value are identical between the
legs for every fixture, mode and operation.

**An independent correctness review** of the diff was run against the worktree,
asked to falsify the identity claim on each of the seven edits. It cleared six
of them — including an exhaustive 134-case differential of the old and new
fixed-width reads over eleven segment shapes, every reachable start position and
`N` in {2, 4}, comparing the value, the resulting `(segment_index, offset)` and
the logical position — and found the formatting-run reservation, which is
recorded under *Validation preserved* with the error-identity consequence it
added. Its three remaining notes — a code comment that elided the reservation,
`frame_header`'s panic *message*, and the debug-build cost of the assertion —
are now in the code comments and in this record.

## Validation preserved

Every structural and limit check survives with its identity: `next_frame` still
refuses a header that would cross the `BoundSheet` boundary, a payload longer
than `MAX_RECORD_BYTES`, a frame that overflows and a scan past
`max_worksheet_scan_records` or `max_worksheet_scan_bytes`, in that order and
with those messages; `parse_globals` still refuses past `max_global_records` and
`max_global_bytes`; `SourceTextSheet::insert` still refuses past
`max_text_bytes` and `max_text_cells`; the SST walk still validates UTF-16
well-formedness through 0576's cold rematerializing path, still checks that
formatting runs lie inside the text and are strictly increasing, and still
parses `ExtRst`.

**One refusal is removed, and it is removed by removing an allocation.** The
source-backed shared-string walk no longer attempts `try_reserve_exact` for
formatting runs, so it can no longer produce
`InvalidData("cannot allocate shared string formatting runs: ...")`. That
reservation guarded a `Vec` the walk **discarded**; at most 65,535 runs × 4
bytes = 256 KiB was ever at stake, nothing is now retained that was not retained
before, and the eager `SharedStringTable`, which keeps its runs, still takes the
reservation and still refuses. This is a bounded allocation that no longer
happens, not a ceiling that was relaxed.

It has a second consequence, which an independent review of this diff named and
which is recorded here rather than left to be discovered. The reservation ran
*before* the first run byte was read, so on a string whose runs are **also**
malformed it used to win: under an exhausted allocator the caller saw the
allocation message, and now it sees whichever run check the bytes actually fail
(`has a formatting run past its text`, `formatting runs are not strictly
increasing`, or an `UnexpectedEndOfStream` from a `Continue` crossing) — or
succeeds, if they are well formed. **Nothing about the input decides this**: it
needs an allocator that refuses 256 KiB, which no test, fixture or default
configuration produces, and the pinned corpus index proves the refusal set over
the 121 fixtures that carry an SST is identical, messages included. It is
reported because "the same refusal set" would otherwise be an overstatement, not
because any reachable input reaches it.

`frame_header`'s panic condition is unchanged (`at + 4 > window.len()`, which
`ensure` makes unreachable); only which index panics first, and therefore the
panic message, would differ in a state that cannot be reached.

Two smaller nuances, for completeness. The `debug_assert_eq!` means
`logical_position` is O(1) in **release** builds only; a debug build still walks
the segments, on purpose, so that every test run checks the invariant. And the
old summation wrapped on overflow in release where the maintained base
saturates — a difference that would need the segment lengths to sum past
`usize::MAX`, which they cannot, because they are lengths of real payload
slices.

## Limitations

- **Nothing is claimed for XLS full text or all cells.** Two of the five
  eager-error sites are in `SourceTextSheet::insert`, on the text-extraction
  path, and `xls_source_attribution` has no full-text or all-cells selector
  (change 0587's measurement blocker). Their effect here is **modelled** from the
  frame-loop sites, not measured.
- **`parse_globals`'s site is measured only indirectly**, through
  `from_shared_ole_file_with_limits` falling 30,734 Ir and the error drop falling
  to zero on the 54016 open.
- **The eager XLS owner is unmeasured.** `SharedStringTable::parse_segments`
  shares `read_u16_continued` and `read_u32_continued` and should benefit from
  their resident fast path, but no eager selector was run in this batch. It does
  **not** benefit from the formatting-run change: it keeps the `Vec`
  instantiation, reservation included.
- **The flagship's wall-clock improvement is not separable from the floor**, as
  stated above. Its counts are.
- **A 1.0% regression is carried on the 54016 open** from an inlining decision
  outside either mechanism (`CellAlignment::parse` outlined from `parse_xf`,
  plus `MeasuredText::consume`). It is inside the saving and is not hidden in a
  mean; a future change to `number_format` may recover it.
- **Every XLS source in this harness is in memory.** On a `from_path` workbook
  each source observation is an `fstat`, and this batch cannot see that cost.
- **Instruction counts rank work, not latency** (change 0579: 1.24% of
  instructions, 6.19% of cycles). Callgrind counts `rep movsb` per byte, so
  every `memcpy` share above is an upper bound.
- **Host quiescence is not established**; the floor is measured, not assumed.
- Withheld entirely: cold-cache, physical-device, range-source, peak-RSS,
  allocation-profile, concurrency-scaling, real-producer and cross-platform
  results.
- **What this does not do:** lazy SST indexing (XLS-2, rank 23) still needs its
  own frozen design record, because it moves *when* a malformed SST is refused;
  the retained sheet index and whole-sheet iterator (XLS-3) still need one; the
  window's zero-fill (0574 opportunity 4) is still blocked by GOAL rule 10; and
  `read_payload`'s remaining 1,456,170 Ir on the 54016 query, plus `query_cell`'s
  2,742,981 and `process_cell`'s 1,636,087, are untouched. Framing overhead on
  that query — `next_frame`, `ensure`/`fill`, `read_payload`, `skip_payload`,
  `query_cell` self and the error drops — falls from **46.3%** of the operation
  (which reproduces change 0587's retained 46.1%) to **38.5%**. It is still the
  largest single term.

## Retained evidence

[`results/change-0595/README.md`](results/change-0595/README.md) — contents
table, provenance, replay instructions, and every raw output this record cites.
