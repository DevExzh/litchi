# 0648: the shared-string resolver reads the string table once, and the XLS whole-sheet walk falls from 16,105 reads to 37

Status: retained, implemented in `litchi-xls` (`resolve_shared_string` and its
`SharedStringResolver`). `performance_claim: none` — this record carries
deterministic read, byte, source-observation and allocation counts over the
whole XLS corpus, change 0627's range-source request counts and sequence
digests, callgrind isolation pairs, `perf stat` cycles and instructions, and
paired medians with a measured A/A floor, and registers no claim-registry
entry.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is the item change [0636](0636-cfb-cursor-bounded-window.md) named as the
next one on the XLS range-source axis, together with survey item **XLS-10**
(`0587-remaining-opportunity-survey.md`): "`resolve_shared_string`'s per-resolve
allocations (chunk `Vec`s, a `slices` `Vec`, a linear segment scan) belong in the
0585 `SharedStringResolver` as a reused scratch buffer and a binary search;
16,055 resolves per `54016` text; unmeasured."

## What was changed

**`crates/litchi-xls/src/workbook/source.rs`.** `SharedStringResolver` — change
[0585](0585-cfb-resumable-cursor-construction.md)'s per-scan state, which until
now held the workbook stream path and one allocation-chain hint — gains a
bounded window over the shared-string table, and `resolve_shared_string_inner`
gains a binary search and two allocation cleanups:

* **The window is all or nothing.** A resolver either retains the string
  table's whole source extent — the `SST` record's payload through the last
  `Continue` payload, headers between them included — or retains nothing and
  reads every entry exactly as before. The extent is computed once, from the
  segment table the open already built, and is taken only when it is at most
  `SHARED_STRING_MAX_WINDOW_BYTES` (256 KiB) and lies inside the declared
  workbook stream; otherwise the window is never opened and nothing about the
  resolve path changes.
* **It is taken late.** A scan reads entries one at a time until it has taken
  more than `SHARED_STRING_WINDOW_RESOLVES` (8) resolves. A selected-cell query
  resolves one shared string and a two-cell query two, so no query pays for a
  window it cannot amortize, and the read set of every small operation is
  unchanged to the byte.
* **A served resolve reads nothing and, in the common case, allocates
  nothing.** An entry inside one `Continue` payload — every entry that does not
  straddle a record boundary — is decoded from a slice of the retained table.
  The `chunks` `Vec`, the zero-filled `Vec` per chunk and the `slices` `Vec`
  that XLS-10 named are all gone from that path; an entry that does straddle a
  boundary keeps one `Vec` of exactly the slices it needs, because the
  `Continue` boundaries are what the BIFF8 string grammar reads.
* **`segments` is searched, not scanned.** `locate_sst_segment` replaces the
  linear walk with `partition_point` over a column the scan builds
  non-decreasing, so locating an entry costs `O(log n)` per resolve instead of
  `O(n)`.
* **The per-entry path reserves what it needs.** When the window does not
  apply, `chunks` is reserved for the segments the entry actually spans rather
  than for every segment from the first one to the end of the table.

Three tests are added to `crates/litchi-xls/tests/source_backed.rs` and one
rewritten, one is added to `crates/litchi-xls/tests/sst_scan_allocations.rs`,
and six unit tests are added to `source.rs`'s own module, with one existing
whole-sheet differential extended to a fixture that reaches the multi-segment
path. No other crate is
touched; `litchi-cfb` needed no accessor.

## Why a retained table and not a sliding window

The brief proposed a bounded window that would "serve most resolves from
retained bytes". The first implementation was exactly that: a 64 KiB window
filled forward from the entry that missed it, growing on change 0568's
512 → 64 KiB schedule. It is implemented, measured and rejected, and the
measurement is the reason this record's design is the shape it is.

The shared-string access order has almost no spatial locality. A temporary
trace of every resolve's `(offset, span)` on three fixtures — retained as
[`trace/`](results/change-0648/trace/) with the simulator that reads it — prices
every policy over the real sequences. On `54016.xls`'s whole-sheet walk, 16,055
resolves over a 225,003-byte table:

The rows below are the simulator's, and it charges one read per resolve; the
resolver itself charges one per **chunk**, which change
[0621](0621-xls-open-fence-count.md) counted at 16,077 for this walk, the extra
22 being entries that straddle a `Continue` boundary. The measured before leg
takes 16,105 reads in total, so the worksheet scan's own window fills are the
remaining 28. The comparison between policies is unaffected.

| policy | reads | bytes |
| --- | ---: | ---: |
| per entry (before) | 16,055 | 323,234 |
| sliding window, 4 KiB | 8,450 | **34,540,295** |
| sliding window, 16 KiB | 7,799 | **125,755,490** |
| sliding window, 64 KiB | 5,539 | **341,656,514** |
| window = the entry's span exactly | 11,906 | 310,560 |
| **the whole table, once** | **1** | **225,003** |

A sliding window at any ceiling below the table's size cuts reads by half and
multiplies bytes by a hundred, because each miss fills a window of which the
average 20-byte string uses a fraction. Reading the table once is better on
**both** axes at the same time, and the reason is the reuse: the walk resolves
16,055 strings out of 7,893 distinct entries, so the per-entry path reads the
repeated ones again every time and ends up reading more bytes than the table
contains.

The 64 KiB sliding window is also what the first implementation shipped, and
`crates/litchi-xls/tests/source_backed.rs::the_whole_sheet_walk_costs_one_scan_not_one_per_cell`
failed it on sight: the walk read **11,251,545 bytes** where a 256-cell
one-at-a-time sample read 10,433,726. That test is the guard that caught it.

## Why it is sound

**No byte outside the string table is ever read.** The window is exactly the
extent the open's own segment table describes, so every byte it reads is a byte
of the `SST`/`Continue` record group the open already read to build the entry
table. The window can read no byte the open did not. A segment table that
reaches past the declared workbook stream length — which the record walk that
built it cannot produce — disables the window outright rather than being clamped
into a region that could leave a late entry outside the retained bytes.

**Bounded memory, and not a new class of it.** One window per resolver, at most
256 KiB, allocated with `try_reserve_exact` and reported as
`SourceBackedError::Allocation { resource: "retained SST window" }` when the
reservation fails, released when the scan ends. The open already holds the whole
workbook-globals substream — this table inside it — in one buffer while it
scans, under a `max_global_bytes` ceiling that is 128 MiB by default and that
can only be larger than the table it contains; and the text projection already
retains up to `max_text_bytes`, also 128 MiB by default. Change
[0576](0576-xls-sst-scan-without-materialization.md)'s rule that the owner retains locators and
not text is untouched: nothing here outlives the resolver.

**Nothing is committed before the step that can fail has succeeded.** A fill
that fails leaves the window empty and gives it up for the rest of the scan, so
no resolve can ever serve bytes that were not completely read.

**The window cannot make an operation fail that would have succeeded without
it.** The two kinds of fill failure part company: an allocator that cannot hand
out the table drops the resolver back to the per-entry path and the resolve
reads its own entry, because that refusal is new in this change and swallowing
it restores exactly what the workbook did before; every other failure — a
mutation, a chain fault, a short read — is the resolve's own and is reported,
and it ends the scan.

**One resolve is never assembled from two reads.** The table is read whole or
not at all, so an entry's bytes always come from one `read_exact` and can never
be spliced from two versions of a file that changed between them.

**Error identity.** Every refusal the per-entry path could report is reported by
the same construction with the same message: `SST not available`,
`Invalid SST index: …`, `SST entry locator has an empty span`,
`SST entry locator is outside its segments`, `SST segment span overflow`,
`SST source offset overflow`, the three `SharedStringScanError` mappings, and the
allocation refusals for `selected SST chunks`, `selected SST entry` and
`selected SST parser segments`. The `Continue` boundaries are preserved exactly —
the slices handed to the parser are the same slices in the same order — which
is what keeps a BIFF8 string that switches its high-byte flag at a record
boundary decoding as it did. The corpus differential below is the evidence:
565 frozen outcomes, including every typed refusal, identical on both legs.

**The observation contract moves exactly as change 0636 moved it.** Change
[0621](0621-xls-open-fence-count.md)'s rule is one observation per read; here a
**fill** is that read. A resolve served from the retained table takes no read and
so no observation, which moves the point at which a mutation is reported from
the resolve after it to the fill after it. Two things bound that, and they are
0636's two:

1. Every operation built on `scan_worksheet` ends with
   `SourceInner::ensure_current`, and `SourceBackedWorkbook::text` ends with one
   too, so the operation-level bracket — the version captured at open, the
   version observed before the result is returned — is exactly what it was. A
   mutation anywhere inside a walk is still refused with the same typed error.
2. A mutation reverted before the next observation is not observed. That is
   already the documented model in `litchi_core::FileVersionPolicy` and on
   `SharedOleStreamCursor::read_exact`. What changes is the width of that
   window, from one resolve to one scan.

`a_mutation_in_any_observation_window_of_a_walk_past_the_threshold_is_refused`
is change 0621's sweep extended to the fill boundary, and it pins both halves.

**ADR reading.** ADR 0003's bounded-resource rule is met by the named ceiling
and the fallible reservation; the window is refused rather than truncated when a
table is larger. ADR 0005's lazy-payload contract is observed: the table is read
on demand, from the same chain walk, by a scan that has shown it will use it,
and no stream is materialized — the window is a clean-value cache of a bounded
region, discarded with the scan. ADR 0006 is untouched: no execution context, no
worker pool, no ambient I/O, no lock. No new `unsafe`, no new dependency, no
public API change and no public leakage of archive types, raw locks or
executors. The cancellation checks are taken in the same places and the same
numbers.

## Measured

Base commit `9f28ea621` (`feat/office-format-completeness`, change 0636), branch
`perf/0648-xls-shared-string-resolver-window`, both legs built
`--release --locked`, every measured process pinned with `taskset -c 23`, other
agents active on the host throughout. Binaries and their SHA-256s are in
[`binaries.sha256`](results/change-0648/binaries.sha256).

### Deterministic counts: every XLS fixture, five operations

126 `.xls` files under `test-data/`; 113 open, giving 565 operation rows. Every
row's outcome is frozen as a digest — the worksheet names, the cell-value list,
the text, the validation report, or the verbatim refusal — and **all 565 are
identical on both legs.**

| operation | reads before | reads after | bytes before | bytes after | observations before | observations after |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `open` | 1,666 | **1,666** | 3,693,129 | **3,693,129** | 1,567 | **1,567** |
| `list` | 1,666 | **1,666** | 3,693,129 | **3,693,129** | 1,673 | **1,673** |
| `validate` | 1,842 | **1,842** | 7,511,960 | **7,511,960** | 1,699 | **1,699** |
| `all-cells` | 26,655 | **959** (−96.40%) | 2,108,161 | **2,023,242** (−4.03%) | 27,019 | **1,323** (−95.10%) |
| `full-text` | 29,708 | **1,226** (−95.87%) | 3,210,676 | **3,065,552** (−4.52%) | 79,061 | **50,579** (−36.03%) |

`open`, `list` and `validate` are identical read for read, byte for byte and
observation for observation across the whole corpus — the resolver is not on
those paths, and the threshold keeps it off the paths that touch it lightly.
**No fixture takes more reads on any operation**, and the corpus reads fewer
bytes overall.

Per fixture:

| fixture | operation | reads | bytes | observations |
| --- | --- | --- | --- | --- |
| `54016.xls` | `all-cells` | 16,105 → **37** (−99.77%) | 938,968 → **840,920** | 16,096 → **28** |
| `54016.xls` | `full-text` | 16,105 → **37** | 938,968 → **840,920** | 36,482 → **20,414** |
| `WithCustomViews.xls` | `all-cells` | 880 → **16** | 146,824 → **141,898** | 883 → **19** |
| `WithCustomViews.xls` | `full-text` | 882 → **18** | 147,652 → **142,726** | 2,103 → **1,239** |
| `ConditionalFormattingSamples.xls` | `full-text` | 341 → **56** | 224,946 → 225,486 | 891 → **606** |
| `59858.xls` | `all-cells` | 29 → **12** | 3,192 → **18,822** | 32 → **15** |

**The honest worst case is `59858.xls`'s whole-sheet walk**, and it is the
largest byte increase anywhere in the corpus: +15,630 bytes for −17 reads. That
sheet resolves just past the threshold on a workbook whose table is 16,384
bytes,
so the table is read whole to serve a handful of entries. At change 0627's
transport the trade is −17 ms of fixed service against +0.15 ms of transfer; on
an owned source it is 15 KiB of extra copying. The largest byte decrease is
−98,048, on `54016.xls`'s walk.

### The harness's own attribution counters, on both source kinds

Change 0605's `xls_source_attribution` binary is an independent instrument with
its own counters and its own oracle, and it measures the **whole lifecycle**
including the open, so its 77 reads are this record's 40-read open plus its
37-read walk. `54016.xls`, worksheet 0, `--all-cells-strategy scan`, three
retained samples:

| leg | mode | operation | reads | bytes | observations | elapsed p50 |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| before | `owned-readat` | `all-cells` | 16,145 | 1,256,139 | 16,117 | 18.176 ms |
| after | `owned-readat` | `all-cells` | **77** | **1,158,091** | **49** | **10.599 ms** |
| before | `file-source` | `all-cells` | 16,145 | 1,256,139 | 16,117 | 24.236 ms |
| after | `file-source` | `all-cells` | **77** | **1,158,091** | **49** | **10.652 ms** |
| before | `owned-readat` | `full-text` | 16,145 | 1,256,139 | 36,503 | 27.364 ms |
| after | `owned-readat` | `full-text` | **77** | **1,158,091** | **20,435** | **16.587 ms** |
| before | `file-source` | `full-text` | 16,145 | 1,256,139 | 36,503 | 33.194 ms |
| after | `file-source` | `full-text` | **77** | **1,158,091** | **20,435** | **19.316 ms** |

Its frozen outcome — the 38,950-cell digest and the 840,804-character text
digest — is identical on all eight rows, and the same binary run over the whole
corpus is a **second, independent differential**: 126 fixtures × two operations
= 252 rows, of which 229 produce a report (the tool declines a fixture with no
worksheet 0, or whose open it refuses), and **all 252 outcome strings are
identical on both legs**. Its aggregates, which include the open the sweep
brackets out:

| operation | rows | reads | bytes | observations |
| --- | ---: | --- | --- | --- |
| `all-cells` | 112 | 23,683 → **2,477** (−89.54%) | 6,116,989 → 6,035,372 (−1.33%) | 23,892 → **2,686** (−88.76%) |
| `full-text` | 117 | 27,311 → **3,052** (−88.83%) | 7,650,840 → 7,506,090 (−1.89%) | 76,890 → **52,631** (−31.55%) |

**No row reads more times**, and the largest per-row byte increase is the same
+15,630 on `59858.xls` that the sweep found, from an instrument that counts it
separately. The `file-source` rows are the same
counters paid as `pread` and `fstat` syscalls; on `all-cells` the gap between
the two source kinds closes from 6.06 ms to 0.05 ms, because 16,145 `pread` and
16,117 `fstat` calls become 77 and 49. Those elapsed figures are three-sample
medians from one short run, not the paired legs below.

### Allocations

The probe's own counting allocator over one complete lifecycle, owned source:

| fixture | operation | allocations | allocated bytes |
| --- | --- | --- | --- |
| `54016.xls` | `all-cells` | 84,182 → **36,042** (−57.19%) | 10,281,754 → **2,144,490** (−79.14%) |
| `54016.xls` | `full-text` | 112,259 → **64,119** (−42.88%) | 22,265,630 → **14,128,366** (−36.55%) |
| `WithCustomViews.xls` | `all-cells` | 7,432 → **4,871** (−34.46%) | 893,206 → **721,648** (−19.21%) |
| `ConditionalFormattingSamples.xls` | `all-cells` | 1,147 → **1,064** (−7.24%) | 712,622 → 715,057 (+0.34%) |
| all three | `open` | identical | identical |

On `54016.xls`'s walk that is **5.24 allocations per resolved shared string
falling to 2.24** — the three `Vec`s XLS-10 named, removed. The remaining 2.24
is the decoded `String` and the walk's own per-cell work, which this change does
not touch.

### Range source

Change 0627's five XLS range-source selectors on `54016.xls`, its transport
(1 ms fixed service per request, 100 MiB/s, 64 KiB maximum range), three
retained samples per case. The before leg reproduces 0627's published counts
exactly — 40, 40, 66, 16,145, 16,145 — which is this packet's provenance check.

| case | requests | bytes | request-sequence SHA-256 | modelled service floor |
| --- | ---: | ---: | --- | ---: |
| `xls_range_source_open` | 40 → 40 | 317,171 → 317,171 | identical | 43.02 ms → 43.02 ms |
| `xls_range_source_open_list_worksheets` | 40 → 40 | 317,171 → 317,171 | identical | 43.02 ms → 43.02 ms |
| `xls_range_source_open_one_cell` | 66 → 66 | 933,004 → 933,004 | identical | 74.90 ms → 74.90 ms |
| `xls_range_source_open_all_cells` | 16,145 → **77** (−99.52%) | 1,256,139 → **1,158,091** | changed | **16,156.99 ms → 88.04 ms** |
| `xls_range_source_open_full_text` | 16,145 → **77** | 1,256,139 → **1,158,091** | changed | **16,156.99 ms → 88.04 ms** |

The three unchanged cases have **identical** digests over the complete ordered
`(offset, requested, returned)` sequence, which is the strongest available
statement that the open, the worksheet listing and a selected-cell query did not
move. The two that changed carry the same frozen observation on both legs.

**The elapsed figure on a range leg is modelled, not measured** — it is
sleep-driven arithmetic over a deterministic request count, as change 0627
established.

### Cycles and instructions

`perf stat -r 5` isolation pairs, N=20 and N=120 complete lifecycles differenced
and divided by 100, owned in-memory source, CPU 23:

| scenario | cycles before | cycles after | Δ | instructions before | instructions after | Δ |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `all-cells` `54016.xls` | 43,916,443 | 14,118,360 | **−67.85%** | 149,789,641 | 47,238,842 | **−68.46%** |
| `all-cells` `WithCustomViews.xls` | 3,084,123 | 1,735,635 | **−43.72%** | 10,944,700 | 6,815,300 | **−37.73%** |
| `full-text` `54016.xls` | 98,656,309 | 61,555,387 | **−37.61%** | 333,341,787 | 230,734,342 | **−30.78%** |
| `all-cells` `ConditionalFormattingSamples.xls` | 589,211 | 428,055 | **−27.35%** | 1,978,425 | 1,510,126 | **−23.67%** |
| `all-cells` `59858.xls` | 700,076 | 568,971 | **−18.73%** | 2,338,829 | 2,130,640 | **−8.90%** |
| `open` `54016.xls` | 817,486 | 818,266 | +0.10% | 3,990,779 | 3,991,674 | +0.02% |

`59858.xls` is the fixture that reads **15,630 more bytes**, and it is still
18.73% cheaper in cycles: seventeen fewer cursor constructions and reads outrun
the extra copy. The `open` control moves by a tenth of a percent on cycles and
two hundredths on instructions, which is the statement that the paths this change does not touch
execute what they executed.

Three `perf stat` windows were run, and
[`cycles/cycles-run-history.txt`](results/change-0648/cycles/cycles-run-history.txt)
keeps all three. The instruction column repeats to within **0.53 points** on
every scenario across all three and to within 0.06 on four of six; the cycle
column repeats to within 3.1 points on five and moves **14.9 points** on
`full-text` `54016.xls` (−28.62%, −43.50%, −37.61%), which is host state on a
machine carrying seven other agents — the same effect change 0636 documented
across its three windows. Every window puts every scenario in the same
direction; the size of each effect is read from the instruction column and the
paired medians below.

Callgrind isolation pairs (N=2 and N=6, differenced and divided by 4) agree,
which is worth saying because change [0604](0604-cfb-append-reads-design.md)
found callgrind overstating `rep movsb` by 35× and this change adds one bulk
copy:

| scenario | before Ir/op | after Ir/op | Δ |
| --- | ---: | ---: | ---: |
| `all-cells` `54016.xls` | 151,483,901 | 50,275,585 | **−66.81%** |
| `all-cells` `WithCustomViews.xls` | 11,213,829 | 7,356,704 | **−34.40%** |
| `full-text` `54016.xls` | 335,421,980 | 234,002,708 | **−30.24%** |
| `all-cells` `ConditionalFormattingSamples.xls` | 3,099,061 | 2,614,050 | **−15.65%** |
| `open` `54016.xls` | 4,685,685 | 4,686,284 | **+0.01%** |

Both instruments put every scenario in the same direction and within six points
of each other, so the copy the window adds is not large enough for the two to
disagree the way 0636's `ConditionalFormattingSamples` validation row did.

### Paired timing

120 samples per leg, order A1 B1 B2 A2, CPU 23, one complete fresh lifecycle per
sample. `owned` is an in-process `ReadAt` over the file's bytes; `file` is
`litchi_core::FileSource`, where every source observation is an `fstat` and every
read a `pread`.

| scenario | A p50 (µs) | B p50 (µs) | Δ p50 | Δ p95 | Δ p99 | A/A floor p50 | B/B p50 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `all-cells` file `54016.xls` | 16,002.9 | 3,185.0 | **−80.10%** | −80.17% | −80.20% | −0.53% | +0.66% |
| `all-cells` owned `54016.xls` | 9,768.4 | 3,161.0 | **−67.64%** | −67.61% | −67.22% | +0.87% | +1.01% |
| `all-cells` owned `WithCustomViews.xls` | 688.6 | 386.9 | **−43.82%** | −43.37% | −43.92% | −0.01% | −0.11% |
| `full-text` owned `54016.xls` | 19,988.0 | 13,013.9 | **−34.89%** | −36.50% | −37.83% | −0.19% | +0.46% |
| `all-cells` owned `ConditionalFormattingSamples.xls` | 128.9 | 92.6 | **−28.15%** | −25.89% | −28.22% | +0.02% | +8.16% |
| `all-cells` owned `59858.xls` | 154.3 | 128.3 | **−16.86%** | −15.62% | −10.29% | +0.17% | +0.44% |
| `open` owned `54016.xls` | 179.5 | 179.1 | −0.22% | −0.24% | −0.01% | −0.05% | +0.48% |

**Every A/A floor in this window is at most 0.87% in absolute value**, and every
effect above is between nineteen and ninety times it. Three windows were run,
for the same reason the cycle pairs were, and
[`timing/timing-run-history.txt`](results/change-0648/timing/timing-run-history.txt)
keeps all three: every scenario repeats to within **3.3 points at p50** across
them and to within 1.9 on six of seven. The one number worth naming is
`ConditionalFormattingSamples.xls`'s **B/B of +8.16%**: that after leg runs in
89–96 µs and its two halves differ by about 7, which is host state on a machine
with other agents on it, not a property of the change — its A/A floor in the
same window is +0.02% and its instruction delta is −23.67%. Nothing here is read
from a single leg.

**No scenario got worse.** Every scenario run for this change is in this table.

## Correctness evidence

**The corpus differential is the oracle.** 565 frozen outcomes over 113
fixtures — the worksheet listing, one worksheet's complete cell-value list, the
workbook text projection, the validation report and every typed refusal —
**identical on both legs**, including `ConditionalFormattingSamples.xls`'s
`source XLS parse error: Invalid record 0x0006: shared Formula metadata requires
a leading PtgExp token`, reproduced verbatim after 341 reads and 224,946 bytes
before and 56 reads and 225,486 bytes after.

**Three new integration tests and one rewritten** in
`crates/litchi-xls/tests/source_backed.rs`:

* `a_text_extraction_observes_the_source_once_per_shared_string_read_and_no_more`
  (rewritten) — change 0621's slope, restated over the range where it still
  applies: four more shared-string cells below the threshold cost four more
  reads and four more observations.
* `a_text_extraction_past_the_resolver_threshold_reads_the_string_table_once` —
  the new slope: twenty-four more string cells past the threshold cost at most
  two more reads, and every added read still carries exactly one added
  observation. The intercept is pinned too — ten resolves cost eight entry reads
  and one fill — so a regression that simply stopped resolving would not pass.
* `a_whole_sheet_walk_reads_the_shared_string_table_once` — `WithCustomViews.xls`
  resolves 862 shared strings in one walk and reads its 101,121-byte table
  **exactly once, in one read**, with fewer than `strings / 10` reads in total.
* `a_mutation_in_any_observation_window_of_a_walk_past_the_threshold_is_refused`
  — change 0621's sweep extended to the fill boundary. The projection is checked
  to cross the fill first; then a mutation is placed after every observation
  window of the operation in turn and the same typed refusal is required from
  all of them, with the control at the end mutating after the last observation
  and requiring success.

**One existing differential extended.**
`the_whole_sheet_walk_agrees_with_selected_cell_queries` now also runs
`WithCustomViews.xls`, and that pairing is the one that reaches the branch this
change makes load-bearing: a walk shares one resolver and crosses the threshold,
so its 862 strings are decoded out of the retained table, while each
selected-cell query builds a fresh resolver and stays on the per-entry path —
and eleven of that workbook's entries straddle a `Continue` boundary across its
thirteen `SST` records, which is where the slice boundaries matter. The
fixtures the test carried before cannot reach it:
`ConditionalFormattingSamples.xls`'s string table is a single 4,599-byte `SST`
record with no `Continue` at all. 3,325 cells compared, no disagreement.

**One new allocation test** in `crates/litchi-xls/tests/sst_scan_allocations.rs`,
under that suite's counting global allocator:
`a_whole_sheet_walk_allocates_less_than_once_per_shared_string_resolved` pins
`54016.xls`'s walk below three allocations per resolved string; it measures
2.24 and measured 5.24 before.

**Six new unit tests** in `crates/litchi-xls/src/workbook/source.rs`:
`the_sst_region_spans_every_segment_and_the_headers_between_them`,
`the_sst_region_is_refused_when_it_does_not_fit_one_window`,
`the_sst_region_never_reaches_past_the_declared_stream`,
`an_empty_or_unrepresentable_sst_region_disables_the_window`,
`the_segment_search_agrees_with_a_linear_walk` (the binary search checked
against the linear walk it replaces, at every logical offset of six segment
tables: an empty one, one segment, three contiguous, and three shapes with empty
`Continue` payloads at the front, in the middle, consecutive, and at the very
end) and `a_locator_outside_every_segment_is_refused_by_name` (both arms: a
locator past a table that has segments, and any locator against a table that has
none).

**An independent review** of the diff ran an exhaustive structural differential
of its own — every accumulation-consistent segment table of one to four segments
with payload lengths in `{0,1,2,3}`, crossed with every entry span, 10,810
`(table, span)` pairs — and found the windowed slice list identical to the
per-entry chunk list on all of them, with no locate mismatch and no out-of-range
`last_segment`. The two coverage gaps and the allocation-failure behaviour it
found are fixed above and in *Why it is sound*; it is cited here because its
harness is not retained, only its result.

**Gates**, all in the worktree, tails in
[`results/change-0648/gates.txt`](results/change-0648/gates.txt):

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-xls --all-targets` | clean, 0 warnings (workspace lints are deny) |
| `cargo doc -p litchi-xls --no-deps` | clean (rustdoc lints are deny) |
| `cargo test -p litchi-xls` | 1,411 passed, 0 failed, 1 ignored across 72 suites |
| `cargo test -p litchi-xlsx -p litchi-xlsb -p litchi` | 2,040 passed, 0 failed, 16 ignored across 103 suites |
| `cargo test -p litchi --features docx,xlsx,pptx,xls` | 266 passed, 0 failed, 7 ignored across 26 suites |
| `cargo test --release --locked` in `tools/perf-baseline` | 531 passed, 0 failed, 1 ignored across 19 suites |

No pre-existing failure was met, so none had to be reproduced on the before
checkout.

## Validation preserved

No validation path was relaxed, bypassed or moved. `validate_source` reads the
same 1,842 requests and 7,511,960 bytes corpus-wide, to the byte. The CFB index
validation, the allocation-chain checks, the directory and input ceilings, the
`max_workbook_stream_bytes` gate, `max_sst_entries`, `max_global_bytes`,
`max_global_records`, `max_worksheet_scan_bytes`, `max_worksheet_scan_records`,
`max_text_cells` and `max_text_bytes` all run exactly where they ran, on exactly
the bytes they ran on. The SST scan that enforces `max_sst_entries` is an
open-time path this change does not touch, and the text budgets are applied by
`SourceTextSheet::insert` after the resolve, unchanged. The BIFF8 string grammar
sees the same segment boundaries, which is the property the `Continue`-straddling
entries depend on.

## Limitations

**What is not claimed.**

* **No claim-registry entry, and no speedup claim beyond these scenarios.** The
  numbers are scoped to seven timing scenarios, six cycle scenarios, five
  callgrind scenarios, five range-source selectors and 113 fixtures on one host,
  one toolchain and one build.
* **The window's failure to allocate is reasoned, not tested.** An allocator
  that cannot hand out the table drops the resolver back to the per-entry path,
  so the refusal this change introduces cannot make an operation fail that would
  otherwise have succeeded. No allocator hook exists in this crate's test
  binaries to force that branch, so it is argued from the code and not
  demonstrated.
* **The range-source figures are modelled, not measured.** They are change
  0627's model arithmetic over a deterministic request count — 1 ms per request
  plus bytes at 100 MiB/s — not observed service.
* **No cold-cache, physical-I/O, network or device result.** The `file` leg runs
  over a warm page cache; what it measures is `pread`/`fstat` call count, not
  physical I/O.
* **A workbook whose string table exceeds 256 KiB gains nothing.** It resolves
  exactly as it did before, and this change is silent about what it would cost
  to serve it. No fixture in this repository's corpus is in that class: 108 of the
  113 fixtures that open carry a string table, in 61 distinct sizes from 8
  bytes to 225,003, and none exceeds the ceiling.
* **The ceiling and the threshold are not swept.** 256 KiB is the smallest
  power of two that covers this corpus's largest string table, and 8 is the
  smallest resolve count that leaves a two-cell query untouched. Neither is
  derived from a cost model, and a workbook in between — a table just under the
  ceiling read by a scan that resolves exactly nine strings — pays the whole
  ceiling once. `59858.xls`'s walk is the corpus's instance of that shape and is
  reported above: +15,630 bytes, −17 reads, −18.73% cycles.
* **The `open` path is unchanged, not improved.** The open still reads and
  scans the whole string table to build the entry locators; this change is
  entirely about what happens afterwards.
* **Peak RSS is not measured.** The window's high-water mark is bounded by the
  ceiling and is reported here as a bound, not as an observation.

**Rejected on the way, and retained.** A sliding 64 KiB window filled forward
from each missing entry, growing on change 0568's schedule, is implemented,
measured and rejected: on `54016.xls`'s walk it reads 341,656,514 bytes where
the per-entry path reads 323,234. The trace of every resolve and the simulator that
prices six ceilings and four policies against it are in the packet; so is the
reason the window here is all or nothing rather than a compromise between them.

**Left open.** The walk's remaining 37 reads on `54016.xls` are the worksheet
substream window of change 0568 and the open; the shared-string path is down to
one. The 2.24 allocations per resolve that remain are the decoded `String` and
the walk's own per-cell work, and are not addressed here.

## Retained evidence

[`results/change-0648/README.md`](results/change-0648/README.md) — the
attribution probe and its manifest template, the two corpus sweeps and their
summary, the allocation table, the resolve traces with the policy simulator that
rejected the sliding window, the callgrind and `perf stat` pairs with their
runners, every timing leg verbatim, the two range-source reports and their
comparison, the gate tails, the decision record and the log paragraphs.
