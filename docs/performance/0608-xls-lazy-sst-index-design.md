# 0608: the XLS shared-string walk is 3.2-35.8% of an open, and every fixture in the corpus needs its last entry

Status: design only. No production change and `performance_claim: none` — the
numbers below are callgrind isolation pairs, native cycle counts, paired medians
with a measured floor and an exact corpus census, reported as evidence and not
registered as claims. **No file under `crates/` was modified by this change.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is item **XLS-2** (rank 23) of change
[0587](0587-remaining-opportunity-survey.md), the frozen design record that
change [0576](0576-xls-sst-scan-without-materialization.md) deferred to and
change [0595](0595-xls-frame-loop-and-sst-walk.md) deferred to again. It
re-measures the walk on the current base, measures for the first time what a
*prefix* index would actually have to walk on real files, writes the design out
in full, and declines to implement it.

## Why this record exists

Change 0576 ended with a sentence that has been carried forward twice:

> **It does not defer SST decoding.** The scan still visits every shared string
> at open time to establish its extent […] A design that indexes lazily […]
> changes *when* a malformed SST is refused, so it needs its own frozen design
> record.

Change 0587 ranked the item on a retained figure — 49.5% of the `54016.xls` open
— and named the two things a design would have to settle: that `segments` is
already free so only `entries` is deferrable, and that deferral moves per-string
refusals from open to the first resolve past the defect, where an
open-and-list-only caller never sees them.

Two facts that decide the item were never measured, and this record measures
both.

**First, the size on the current base.** 0595 landed after the survey and made
the walk cheaper. The 49.5% the queue ranks is a pre-0595 number.

**Second, and decisively: a deferred index is a *prefix* index.** Entry *k*'s
extent is known only after entries 0..*k*−1 have been framed, because a BIFF8
shared string is self-delimiting only by walking it. So the question is not "how
many strings does a scenario resolve?" but "what is the **highest index** it
resolves?" — and nobody had counted that. This record counts it across every
`.xls` and `.xlt` fixture in the repository.

## What the walk costs on this base

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Every measured child is `tools/perf-baseline`'s
`xls_source_attribution` built `--release --locked`, pinned to CPU 28 with
`setarch x86_64 -R` (ASLR off) and single-threaded. **Host quiescence is not
established**: eight measurement agents shared the machine and the load average
ran between 6.1 and 10.8 across the capture window. That is why every floor
below is measured rather than assumed. Binary hashes and fixture hashes are in
[`environment.json`](results/change-0608/environment.json).

Three legs were built from this worktree, differing **only** in
`crates/litchi-xls/src/records.rs`. Two of them are **measurement scaffolds and
not candidates**; they exist to split one number into two, and their patches are
retained in the packet:

| leg | what it does |
| --- | --- |
| `base` | the unmodified base commit |
| `nostore` | walks every shared string exactly as production does, but reserves nothing and records nothing |
| `nowalk` | builds `segments`, runs every SST *header* check, then skips the per-string walk entirely |

`base − nowalk` is the ceiling of fully deferred indexing at open.
`base − nostore` is the cost of the storage alone — the two `logical_position()`
calls, the `entries.push` and the one `try_reserve_exact` per open.

### The control: not one byte of I/O moves, and nothing outside the SST moves

Eighteen counter cells per leg — three fixtures × two in-memory source modes ×
`open`/`list`/`one-cell` — in
[`counters.txt`](results/change-0608/counters.txt). Every `open` and every `list`
cell is **identical across all three legs**:

| fixture | reads | read bytes | `version()` | `len()` | seeks |
| --- | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 53 | 565,201 | 29 | 1 | 0 |
| `WithCustomViews.xls` | 16 | 110,242 | 22 | 1 | 0 |
| `54016.xls` | 40 | 317,171 | 25 | 1 | 0 |

Those rows reproduce 0595's, 0576's, 0574's and 0565's retained figures exactly.
**This is the single most important fact for the design**: the SST payload bytes
are read at open under *every* index shape, because `parse_globals` frames every
globals record through `GlobalsBuffer` whether or not the SST is walked.
Deferral is therefore a CPU question at open, not an I/O one — and, as §"Where
the deferred index lives" shows, it becomes an I/O question on the *resolve*
path instead.

The per-symbol tables say the same thing from the instruction side: on all three
fixtures every symbol outside the SST walk is identical to within a handful of
instructions across the three legs (`__memcpy_avx_unaligned_erms` 675,629 /
675,625 / 675,592 on the flagship; `collect_exact`, `next_chain_sector`,
`validate_stream_allocations`, `parse_xf` byte-for-byte equal).

### Instructions per operation, callgrind isolation pairs

Change 0574's method, as 0576, 0584 and 0595 used it: a large-sample and a
small-sample child are differenced and divided by the extra operations, so
everything that runs once per child cancels. Both a self-cost and an
`--inclusive=yes` annotation are taken from the same raw profile.

| fixture | operation | operation Ir | `scan_shared_string_records` inclusive | share |
| --- | --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | open | 2,326,958 | 74,236 | **3.19%** |
| `WithCustomViews.xls` | open | 808,275 | 147,687 | **18.27%** |
| `54016.xls` | open | 5,335,047 | 1,906,053 | **35.73%** |
| `ConditionalFormattingSamples.xls` | one-cell | 2,548,648 | 74,268 | 2.91% |
| `WithCustomViews.xls` | one-cell | 827,802 | 147,687 | 17.84% |
| `54016.xls` | one-cell | 18,603,947 | 1,906,053 | 10.25% |

The `open` column reproduces 0595's after-leg totals to 0.04%, 0.04% and 0.76%.
The flagship and `WithCustomViews` agree almost exactly; `54016.xls` differs by
0.76% because change 0589 (`perf(cfb): hash an empty overlay once, not twice`)
landed between 0595's base and this one and is the only `crates/` change on this
path since.

**The survey's 49.5% is now 35.73%.** Per shared string the walk is 241.5 Ir on
`54016.xls` (1,906,053 / 7,893), 311.6 Ir on `WithCustomViews.xls` and 245.0 Ir
on the flagship — against the 454 Ir per string change 0587 retained for the
pre-0595 scan, which is 0595's own claim measured from a third direction.

Splitting that with the two scaffolds:

| fixture | open Ir | `base − nowalk` (defer everything) | `base − nostore` (defer the storage only) |
| --- | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 2,326,958 | **74,108 (3.18%)** | 6,950 (0.30%) |
| `WithCustomViews.xls` | 808,275 | **146,246 (18.09%)** | 9,884 (1.22%) |
| `54016.xls` | 5,335,047 | **1,907,192 (35.75%)** | 111,236 (2.09%) |

Self costs on `54016.xls`, base → `nostore` → `nowalk`: `MeasuredText::consume`
759,098 → 759,098 → **0**; `walk_one_shared_string` 742,217 → 742,217 → **0**;
`SstCursor::walk_formatting_runs` 165,818 → 165,818 → **0**;
`scan_shared_string_records` self 238,178 → 80,293 → 1,746. The scaffolds do
exactly what they say.

One callgrind artifact is named rather than rounded away. On `54016.xls` the
`nostore` leg's `__memcpy_avx_unaligned_erms` **rises** from 1,069,111 to
1,116,137 Ir, +47,026: dropping the `entries` reservation changes the heap
layout and some later growth copies more. Callgrind charges `rep movsb` once per
byte, so that figure is an upper bound, and it is exactly why the native
instruction count below puts the same saving at 158,478 rather than 111,236 —
111,236 + 47,026 = 158,262, a 0.14% reconciliation between the two tools.

### Cycles and IPC, native `perf stat`, same isolation method

Change 0579 established that instruction share mis-ranks pointer-chase work, and
0595 showed callgrind and `perf` disagree in level on this very path. Each
isolation-pair leg is the median of five repetitions.

| fixture | open cycles | `base − nowalk` | | `base − nostore` | | IPC base → nowalk |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xls` | 345,078 | 10,302 | **2.99%** | 1,504 | 0.44% | 3.41 → 3.30 |
| `WithCustomViews.xls` | 128,713 | 23,577 | **18.32%** | 4,025 | 3.13% | 3.98 → 3.48 |
| `54016.xls` | 874,476 | 271,026 | **30.99%** | 26,989 | 3.09% | 4.61 → 3.52 |

Native instructions per open move by 6.26% / 28.47% / 47.33% (`nowalk`) and
0.56% / 1.96% / 3.93% (`nostore`).

**The A/A floor, measured in the same window.** Two extra legs were built by
copying the base binary and run through the identical pipeline. On the
isolation-pair cycle metric the six same-binary comparisons span **−0.17% to
+0.66%**; on native instructions they span **−0.10% to +0.00%**, which is the
evidence that the isolation-pair method itself is sound here.

### Wall clock, paired, A1 B1 B2 A2

`p50` of the measured `open` in nanoseconds, 400 samples per round after 50
warmups, four rounds in the order base, scaffold, scaffold, base. `dir1` is
`b1` against `a1` and `dir2` is `b2` against `a2`; `A/A` and `B/B` are the same
binary against itself in the same window. The two p50 columns show `a1` and
`b2`, so neither percentage is the ratio of the two columns beside it — each is
computed from its own adjacent pair, and all four rounds are in
[`analysis.txt`](results/change-0608/analysis.txt).

| scaffold | fixture | mode | base p50 (`a1`) | scaffold p50 (`b2`) | dir1 | dir2 | A/A | B/B |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `nowalk` | `54016` | owned | 191,616 | **131,951** | **−31.10%** | **−31.65%** | +0.75% | −0.06% |
| `nowalk` | `54016` | file | 207,931 | **148,530** | **−28.52%** | **−29.76%** | +1.70% | −0.07% |
| `nowalk` | `WithCustomViews` | owned | 27,340 | **21,730** | **−20.08%** | **−20.43%** | −0.11% | −0.55% |
| `nowalk` | `WithCustomViews` | file | 35,440 | **29,340** | **−16.21%** | **−15.74%** | −1.75% | −1.20% |
| `nowalk` | `Conditional…` | owned | 71,080 | 68,940 | −2.57% | −3.47% | +0.48% | −0.45% |
| `nowalk` | `Conditional…` | file | 89,970 | 86,940 | −3.52% | −3.56% | +0.20% | +0.16% |
| `nostore` | `54016` | owned | 191,466 | 184,731 | −3.60% | −3.66% | +0.15% | +0.08% |
| `nostore` | `54016` | file | 208,406 | 201,726 | −3.14% | −3.36% | +0.16% | −0.07% |
| `nostore` | `WithCustomViews` | owned | 27,340 | 26,446 | −4.06% | −2.45% | −0.84% | +0.82% |
| `nostore` | `WithCustomViews` | file | 34,820 | 34,210 | −0.83% | −2.92% | +1.21% | −0.93% |
| `nostore` | `Conditional…` | owned | 71,346 | 71,550 | **+0.09%** | **+1.39%** | −1.09% | +0.20% |
| `nostore` | `Conditional…` | file | 90,636 | 90,346 | −1.10% | +0.59% | −0.90% | +0.79% |

**The measured floor across all 24 same-binary comparisons is −1.75% to
+1.70%.** The `nowalk` rows on `54016.xls` and `WithCustomViews.xls` are 9 to 18
times that floor and their two directions agree to within 1.3 percentage points;
the flagship's are about twice it. The `nostore` rows are a different story: only
`54016.xls` is separable, and on the flagship the scaffold is *slower* in both
directions of one mode.

## What a prefix index would have to walk

This is the measurement the design turns on, and it did not exist. The probe
([`probe/sst_prefix_census.py`](results/change-0608/probe/sst_prefix_census.py))
walks the CFB container, the BIFF8 framing, the SST header and every `LabelSst`
record of each fixture with pure arithmetic. It shares no code with `litchi-xls`,
so it is an independent oracle; its CFB/BIFF walk is adapted from change 0584's
`sst_walk.py`, which answered a different question with the same framing. Two
fidelity points it had to get right, and 0584's probe does not: a `Workbook`
stream shorter than the header's 4,096-byte cutoff lives in **mini sectors**
inside the root entry's stream, and seven fixtures carry a legacy BIFF5 `Book`
entry *before* the BIFF8 `Workbook` in directory order, so the probe follows
`select_workbook_stream`'s preference — `Workbook`, then `Book` — rather than
taking whichever comes first. Re-running the retained probe with those two
corrections disabled turns 3 skipped fixtures into **40**; with only the
mini-stream read disabled, 35; with only the stream preference disabled, 34.

Full output: [`sst-prefix-census.jsonl`](results/change-0608/sst-prefix-census.jsonl);
folded tables: [`census-summary.txt`](results/change-0608/census-summary.txt).

| | count |
| --- | ---: |
| `.xls` and `.xlt` fixtures under `test-data` | 126 |
| carrying no SST record | 3 |
| carrying an SST record | 123 |
| refused at open by an SST **header** check | 4 |
| indexed at open today | 119 |
| declaring at least one shared string | 94 |
| of those, carrying at least one `LabelSst` cell | **94** |
| shared-string entries over the corpus | **17,434** |
| `LabelSst` cells over the corpus | 30,316 |

**The probe reproduces change 0576's corpus totals exactly, from code that
shares nothing with the scan.** 0576 counted 126 fixtures, 2 refused by
`litchi-cfb` before any SST is reachable, 3 carrying no SST, 121 carrying one,
117 indexed, 4 refused, and **17,434** entries compared. This probe does not
apply `litchi-cfb`'s FAT validation, so it reads the two `litchi-cfb` refuses —
`1900DateWindowing.xls` and `1904DateWindowing.xls` — and every other row falls
out: 123 = 121 + 2, 119 = 117 + 2, the same 3 carry no SST, the same 4 are
refused, and the entry count agrees to the unit.

The four refused are `password.xls`, `35897-type4.xls` and
`xor-encryption-abc.xls` — encrypted, so their count words are ciphertext and
fail the signed-count check — and `57456.xls`, which declares 1,761 unique
strings and 0 total. They are exactly the four change 0576's differential
refuses, identified here independently and with the reason each fails.

**The decisive count: on 94 of 94 fixtures the highest SST index any cell
references is the last entry.** `idx_max == cstUnique − 1` everywhere. A full
text, or an all-cells read, therefore drives a prefix index to completion on
every file in this corpus and saves **exactly nothing** — while paying for the
deferral machinery and, under the design below, for a second read pass over the
SST.

What a prefix index walks, as a share of the table:

| scenario | min | p25 | median | p75 | max |
| --- | ---: | ---: | ---: | ---: | ---: |
| open, or list | — | — | **0%** | — | — |
| one uniformly chosen string cell | 35.34% | 55.04% | **63.88%** | 100.00% | 100.00% |
| the first string cell in stream order | 0.17% | 20.00% | 50.00% | 100.00% | 100.00% |
| a full text, or all cells | — | — | **100.00%** | — | — |

The three profiled fixtures, and the cell the harness's `one-cell` selector
actually reads (row 1, column 0 of the profiled worksheet):

| fixture | unique | SST bytes | segments | `LabelSst` cells | `idx_max` | mean prefix | harness cell |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| `ConditionalFormattingSamples.xls` | 303 | 4,599 | 1 | 658 | 302 | 131.5 | `isst` 3 |
| `WithCustomViews.xls` | 474 | 101,073 | 13 | 862 | 473 | 266.6 | not a string |
| `54016.xls` | 7,893 | 224,895 | 28 | 16,055 | 7,892 | 2,789.3 | not a string |

The counters confirm the probe against production from the other side: the base
leg's flagship `one-cell` returns `string:4:Date` and issues one more positional
read (61 against 60), seven more bytes and four more `version()` observations
than a leg with no index; the scaffold legs return
`Invalid SST index: 3 (max: 0)` — naming the same `isst` 3 the probe predicted.
`54016.xls`'s and `WithCustomViews.xls`'s harness cells resolve no string at all
and their counters are identical on every leg.

So the honest statement of the item's reach is:

- **open, and list**: the full 3.19% / 18.27% / 35.73% (Ir) or 2.99% / 18.32% /
  30.99% (cycles).
- **one cell**: everything, when the cell is not a shared string — which is 2 of
  the 3 profiled cells; `(isst + 1) / unique` of the walk otherwise, 1.3% on the
  flagship's cell, a 63.88% median over the corpus for a cell picked at random.
- **full text, all cells**: **zero**, on all 94 fixtures, plus the design's
  costs.

## The frozen design

### D1. What is deferrable

`SharedStringSstScan { segments, entries }`
(`crates/litchi-xls/src/records.rs`). `segments` is built from the framed
SST/`Continue` records *before* any string is walked — one push per record, 28 on
`54016.xls`, 13 on `WithCustomViews.xls`, 1 on the flagship — and is already
free. Only `entries` is deferrable, and it is a prefix: `entries[k].end ==
entries[k+1].start` holds by construction, because `start` is
`cursor.logical_position()` immediately before `walk_one_shared_string` and
`end` is the same call immediately after, with nothing consumed between
iterations.

### D2. Where the deferred index lives, and what it must re-read

The walk runs over an `SstCursor<'_>` borrowing `&[&[u8]]` payload slices of the
framed globals records, which in turn borrow a `Vec<u8>` moved out of
`GlobalsBuffer`, truncated to the globals length and **dropped when
`parse_globals` returns**. `ParsedGlobals` carries offsets, never bytes, and
`SourceInner` retains only `Arc<SharedStringSstScan>`. A deferred
walk therefore cannot resume over the same memory. It has two shapes:

- **(A) re-read the SST region from the CFB** at first use, through
  `stream_cursor_at_hinted`, exactly as `resolve_shared_string` already does per
  entry; or
- **(B) retain the SST payload bytes** at open — 224,895 B on `54016.xls`,
  101,073 B on `WithCustomViews.xls`, 4,599 B on the flagship.

The counters above prove the bytes are read at open under both, so (B) costs a
copy and a retention rather than a read. **(B) is rejected**: change 0300 fixed
the retained shape deliberately — "does not retain … a complete raw SST copy" —
and reviving it would put a `max_global_bytes`-bounded (128 MiB by default)
buffer back into `SourceInner` to save instructions. The design takes **(A)**,
and accepts that it moves cost onto the resolve path.

The state:

```rust
pub(crate) struct DeferredEntries {
    /// `cstUnique`, already checked against `max_sst_entries` at open.
    unique: usize,
    /// Logical length of the SST payloads; the walk may never pass it.
    logical_len: usize,
    state: Mutex<PrefixState>,
}

struct PrefixState {
    /// A strict prefix: `known[k]` is entry `k` for every `k < known.len()`.
    known: Vec<SharedStringEntryLocation>,
    /// The first refusal the walk produced, memoised.
    refusal: Option<StoredRefusal>,
}
```

`SourceInner` keeps `Arc<SharedStringSstScan>`; only the `entries` field's type
changes.

**`StoredRefusal` is load-bearing, not bookkeeping.** `SourceBackedError` is not
`Clone` — it carries `String` and `litchi_biff::Error` — so the design stores the
owned payload fields and rebuilds the identical value. Without memoisation, a
second resolve past a defect would re-read the source and could report
`SourceChanged` instead of the parse refusal, so the same call would report two
different errors on two invocations. That is a contract break, and the
memoisation is what prevents it.

### D3. The extension operation

```rust
fn extend_to(
    &self,
    owner: &SourceInner,
    wanted: usize,
    execution: Option<&ExecutionContext>,
    strings: &mut SharedStringResolver<'_>,
) -> Result<SharedStringEntryLocation>
```

In this order, and the order is the contract:

1. `wanted >= unique` → today's `CellValue::Error("Invalid SST index: {i} (max:
   {n})")`, with **no walk and no lock**. Change 0300's invalid-index text stays
   exactly where it is and stays O(1).
2. Acquire the lock, copy out `known[wanted]` if present, release. A hit is one
   uncontended acquisition and a 16-byte copy.
3. On a stored refusal, rebuild and return it. Release the lock first.
4. Outside the lock, read forward from `known.last().end` (or 8, the SST header
   length, when empty) and walk entries `known.len() ..= wanted` with
   `walk_one_shared_string::<MeasuredText>` — **the production walk, one
   implementation, not a second copy**, exactly as 0576 and 0595 kept it.
5. Re-acquire the lock and merge. If another thread has already extended past
   `wanted`, discard this walk and take theirs; the walk is a pure function of
   the bytes, so first-writer and last-writer agree by construction.

**The lock is never held across a source read.** A cancellation or a
`SourceChanged` on one thread therefore cannot block another, and poisoning is
handled the way `litchi-xls` already handles it in
`SourceCheckedTextSink::record_failure` —
`.unwrap_or_else(|poisoned| poisoned.into_inner())`. The cost is duplicated walk
work under contention, bounded by the number of concurrent resolvers.

### D4. Re-reading the bytes

The walk needs logically contiguous bytes across `Continue` boundaries.
`resolve_shared_string` already builds that shape: per-segment chunks in a
`Vec<Vec<u8>>`, slices in a `Vec<&[u8]>`, fed to `SstCursor::new`. The extension
needs the same over `[known_end, end_of(wanted))` — whose end is *unknown before
the walk*, which is the whole difficulty. The design reads a window that doubles,
capped at `logical_len`, mirroring `GlobalsBuffer`'s own shape, with
`try_reserve` before every fill and `owner.ensure_current()` before every fill,
as `resolve_shared_string` fences per chunk today.

Cost on the resolve path, stated up front rather than discovered later: the
prefix bytes are read a **second** time, once in total as the index grows. On
`54016.xls` a full text would read at most 224,895 B on top of the 317,171 B the
open already reads — **+70.9% of the open's read bytes** — to save 35.75% of the
open's instructions, on a scenario the census says saves nothing. On a range
source that is a straight regression, which is the concern changes 0572 and 0577
exist to protect.

### D5. Which refusals move, and to where

**Stay at open** — every check that runs before the per-string loop, in
`parse_globals` and in the pre-loop part of `scan_shared_string_records`:

- in `parse_globals` (`workbook/source.rs`):
  `"Workbook globals contain multiple SST records"`;
  `Allocation { resource: "SST record references", requested: 1 }`;
  `InvalidLength { expected: 8 }` for an SST payload shorter than its header;
  and `ResourceLimit { resource: "SST entries", .. }` from `max_sst_entries`;
- in `scan_shared_string_records` before the loop (`records.rs`):
  `UnexpectedRecordType` for a first record that is not `SST` (0x00FC) or a
  non-`Continue` (0x003C) after it;
  `Allocation { resource: "SST segment locator" | "SST parser segments" }`;
  `UnexpectedEndOfStream("SST header")`;
  `"SST payload length overflow"`;
  `"SST counts must be non-negative signed integers"`;
  `"SST total count is smaller than its unique count"`;
  `"SST unique count does not fit in usize"`;
  and `"SST declares {unique_count} strings but its records are too short"`.

**Move to the first resolve at or past the defect** — every refusal
`walk_one_shared_string` can raise:

- `Error::UnexpectedEndOfStream` with each of its per-field contexts —
  `shared string header`, `shared string character count`,
  `shared string flags`, `shared string rich-text count`,
  `shared string extension length`, `continued shared string character data`,
  `shared string continuation flags`, `shared string formatting run`,
  `shared string ExtRst`;
- `"cannot allocate shared string ExtRst: …"`, `read_bytes`'s `try_reserve_exact`
  refusal for the phonetic payload — the one allocation refusal the walk still
  takes after change 0595 removed the formatting-run reservation;
- `"shared string {string_index} has a negative extension length"`;
- `Encoding error: UTF-16 decoding error: invalid utf-16: lone surrogate found`,
  which 0576's cold rematerialising path produces and which must keep producing
  it;
- `"a UTF-16 shared string is split inside a code unit"` and
  `"invalid shared string continuation flags 0x{continuation_flags:02X}"`, both
  live; and `"shared string character data does not end at a record boundary"`,
  which 0576 proved unreachable before and after and which is listed for
  completeness rather than as a live refusal;
- `"shared string {string_index} has a formatting run past its text"` and
  `"shared string {string_index} formatting runs are not strictly increasing"`;
- **every refusal `parse_phonetic_string` raises**, reached from the walk when
  the string carries an `ExtRst` block. There are at least nine and the record
  does not enumerate them as a closed set, because the function is a second
  parser inside the walk:
  `InvalidLength { expected: 14 }` and `InvalidLength { expected: required }`;
  `"shared string {string_index} has invalid ExtRst string counts"`;
  `"ExtRst text length overflow"`, `"ExtRst run length overflow"` and
  `"ExtRst length overflow"`; `Encoding("ExtRst UTF-16 decoding error: …")`;
  `"shared string {string_index} has an invalid ExtRst phonetic run"`;
  `"shared string {string_index} ExtRst runs exceed the base string"`; plus
  every `binary::read_u16_le` short-read inside it;
- `Allocation { resource: "SST entry locator" }`, if the reservation is deferred
  with the index. It is the one entry in this list that is **not** raised inside
  `walk_one_shared_string` today — it guards the `entries` reservation before the
  loop — and D6 keeps it at open, so it moves only under a variant this design
  rejects.

**Two refusals appear that cannot occur today**, because today the walk runs over
bytes already resident:

- `SourceChanged`, if the artifact is mutated between open and the first
  resolve. ADR 0005 permits it — "mutation during a read returns
  `SourceChanged`" — but it is a refusal at a place that has none today, and it
  is the reason D2's memoisation exists.
- `ExecutionContext` cancellation inside the index build.

**Error identity.** The typed value is unchanged: `map_shared_string_error`
already maps `SharedStringScanError::{Biff, Invalid, Allocation}` to
`SourceBackedError::{Parse, Parse, Allocation}`, and `resolve_shared_string`
already returns `Result<CellValue, SourceBackedError>`, so every moved message
is byte-identical in its new position. What changes is **which call reports it**:
`Workbook::open` and `worksheet_names()` now succeed on a workbook whose SST is
malformed past the header. That is change 0576's error-identity trap, and it is
the reason this item needs a record rather than a patch.

**The corpus cannot exercise any of it.** Measured above: of the 123 fixtures
carrying an SST, four are refused by header checks and **not one is refused by
the per-string walk**. Change 0576's differential says the same from the other
side — 117 of 121 indexed identically, 4 refused identically. Every refusal this
design moves fires on **zero** of the repository's fixtures, so the differential
that would prove the move safe has to be built from synthetic malformed SSTs that
do not exist. That is the strongest single argument for freezing rather than
implementing: the change cannot be validated against the corpus it would ship
against.

### D6. Budget accounting, and why deferral saves no memory once a string is read

`max_sst_entries` still bounds `cstUnique` at open; nothing is relaxed, and the
check stays where it is (`workbook/source.rs`, before the scan).

The retained weight becomes variable: `16 × known.len()` instead of
`16 × unique`. Three consequences, and the third is the one that matters.

1. The index is a **monotone prefix, not a cache**. ADR 0005's "clean parsed
   values are evictable" would permit dropping it and re-walking; the design
   forbids that, because a second walk of a defective table could report a
   different error from the memoised one. The index is pinned by the snapshot for
   the snapshot's life. That is a narrower reading of ADR 0005 than the clause
   allows, and it is deliberate.
2. Every growth goes through `try_reserve`, so
   `Allocation { resource: "SST entry locator", requested }` keeps its identity.
3. **But `requested` is part of the message.** Preserving it means reserving
   `unique` — one `try_reserve_exact(unique)` — on the *first* extension rather
   than in geometric steps. So from the first resolved string onward the retained
   weight is exactly what it is today, and **the memory saving exists only for
   callers that never resolve a shared string**. Deferral is an instruction
   saving for open-and-list, not a retention saving for anything else.

### D7. How `SharedStringResolver` (0585) consumes it

`resolve_shared_string` today is: the `segments.is_empty()` check that returns
`CellValue::Error("SST not available")` → index bounds check →
`entries.get(index)` → empty-span check → fence → linear segment scan → cursor →
chunks → decode. Under the design only `entries.get(index)` changes, to
`extend_to(index)`.

The resolver already owns the SST region's `StreamChainHint` (change 0585). The
extension walk moves strictly forward through the SST and the decode then reads
inside the entry it has just found, so the two share one hint naturally and in
the right order — unlike the worksheet cursor, which 0585 showed thrashes against
a shared hint. But `StreamChainHint<'a>` borrows `&'a SharedOleFile` and cannot
be stored in `SourceInner` beside the index, so `extend_to` must take
`&mut SharedStringResolver<'_>`, which is the signature `resolve_shared_string`
already has. **The index is shared state; the chain hint is per-scan state; the
design keeps that split**, and it is the reason `extend_to` is a method taking
the resolver rather than a method on the workbook.

One caution inherited from 0585: a full text resolves in worksheet order, not
index order. This record's census does not measure monotonicity; change 0584's
retained `sst-corpus-full.txt` does, and finds backward steps on real files —
222 of 1,508 resolves on `FormulaEvalTestData.xls`, 14.7%. The
extension walk *is* monotone, so once the prefix is complete the hint behaves
exactly as it does today; before that, a backward resolve discards the hint the
forward walk left, exactly as 0585 described.

### D8. Admission gates

No part of this design may land without all of these.

- **Byte-identical `entries`.** Change 0595's
  `the_sst_index_over_the_corpus_is_pinned` — 121 fixtures, digest
  `0x9cb14f5daa02eebc` — must reproduce exactly when the index is driven to
  completion, per fixture and over the corpus, re-pointed at a "resolve every
  entry" driver instead of at the open.
- **Change 0576's differential, unchanged and still green**:
  `every_sst_fixture_indexes_identically_both_ways`, 121 fixtures, 17,434
  entries, 4 identical refusals.
- **A synthetic malformed-SST differential, which does not exist today.** For
  each reachable refusal listed in D5 — twenty-seven named, one of them
  unreachable, plus the `binary::read_u16_le` short reads inside
  `parse_phonetic_string`, which are not a closed set — a fixture whose SST is
  malformed at entry *k*, asserting that (a) `open` and `worksheet_names()` now succeed, (b)
  resolving entry *k*−1 succeeds, (c) resolving entry *k* reports the
  byte-identical message today's open reports, (d) resolving entry *k* a second
  time reports the same message and issues **no further reads**, (e) an
  out-of-range index still returns `CellValue::Error("Invalid SST index: …")`
  and never walks.
- **Counters.** `open` and `list` reads, bytes and `version()` calls must be
  unchanged — the control, whose values are retained here (53/565,201/29,
  16/110,242/22, 40/317,171/25). The `one-cell` and text paths must have their
  **new** read counts reported, not folded into a mean.
- **Concurrency.** Two threads resolving disjoint indices on one
  `SourceBackedWorkbook` must return identical values, must not deadlock, and
  must survive a poisoned lock the way the existing sites do.
- **Cycles, not instructions**, on the three fixtures plus an A/A in the same
  window, and on the text and one-cell paths as well as the open — because those
  paths lose.
- The crate gates: `cargo fmt --all --check`, `cargo clippy -p litchi-xls
  --all-targets`, `cargo test -p litchi-xls`, `cargo doc -p litchi-xls
  --no-deps`.

### D9. Falsification

- **F1 — falsified as a saving for the corpus's dominant scenario, and this is
  measured, not predicted.** `idx_max == unique − 1` on 94 of 94 fixtures, so a
  full text or an all-cells read drives the prefix to completion and saves zero,
  while paying D4's second read pass. Change 0587 wrote "a saving only when few
  strings are resolved"; the corpus says *no* fixture resolves few.
- **F2 — falsified as a memory saving** for any caller that reads one string, by
  D6.3.
- **F3 — the remaining scenario is open-and-list**, where the saving is 2.99% to
  30.99% of the open in cycles. That is not falsified. It is also the scenario
  with the least reach: the facade routes `.xls` opens to this owner and every
  read path beyond listing resolves strings.
- **F4 — the safety argument rests on a differential that does not exist.** No
  fixture in the repository is refused by the per-string walk, so nothing in the
  corpus can detect the refusal move.

## The one form in which no refusal moves, and what it saves

Change 0608's brief asked whether a form exists in which *no* refusal moves —
for example, keep validating every string at open and store no entries until
first use — and whether it saves anything.

**The form exists, and it is exactly the `nostore` scaffold**: walk every string
at open, so every per-string refusal fires at open with today's message in
today's position, and rebuild the extents on first use. Measured on this base it
saves:

| fixture | callgrind Ir | native Ir | native cycles | wall clock p50, both directions |
| --- | ---: | ---: | ---: | --- |
| `ConditionalFormattingSamples.xls` | 6,950 (0.30%) | 6,598 (0.56%) | 1,504 (0.44%) | +0.09% / +1.39% owned, −1.10% / +0.59% file |
| `WithCustomViews.xls` | 9,884 (1.22%) | 10,018 (1.96%) | 4,025 (3.13%) | −4.06% / −2.45% owned, −0.83% / −2.92% file |
| `54016.xls` | 111,236 (2.09%) | 158,478 (3.93%) | 26,989 (3.09%) | −3.60% / −3.66% owned, −3.14% / −3.36% file |

against a measured wall-clock floor of −1.75% to +1.70% and a measured
isolation-pair cycle floor of −0.17% to +0.66%.

Three things disqualify it.

1. **It is an open-and-list saving only, and it makes every other scenario
   worse.** Rebuilding the extents on first use is the full walk again — 3.18% to
   35.75% of an open in instructions — *plus* D4's re-read of the SST bytes,
   which today do not have to be read a second time at all. A caller that reads
   one string pays roughly twice the walk instead of once.
2. **The measured saving is small and, on the flagship, not separable.** Two of
   six wall-clock cells are positive; the flagship's own A/A excursion (0.45%
   cycles) is as large as its saving (0.44%).
3. **The scaffold does not measure exactly this form.** `nostore` also drops the
   `try_reserve_exact(unique)`, and keeping "no refusal moves" literally means
   keeping that reservation at open. The difference is one allocation per open,
   which is why the numbers above are the form's *upper* bound rather than its
   value — and it cuts both ways, since dropping the reservation is also what
   produced the 47,026 Ir `memcpy` artifact on `54016.xls`. By D6.3, keeping the
   reservation is also what removes the memory argument entirely.

`docs/GOAL.md`'s decision rules say to keep only what is statistically and
practically useful and to revert speculative complexity. A 0.44–3.13% cycle
saving on open-and-list, bought by making every string-reading scenario pay the
walk twice, is neither. **It is not implemented.**

## Why nothing was implemented

The item is declined on three measured grounds, in this order:

1. Its headline reach is zero on the scenario the corpus is made of: 94 of 94
   fixtures drive a prefix index to its last entry (F1).
2. Its memory argument does not survive preserving the `Allocation` message
   (F2, D6.3).
3. Its safety argument cannot be validated by any fixture in the repository
   (F4, D5).

What remains — 2.99% to 30.99% of an open in cycles for a caller that opens and
lists and never reads a string — is real, is measured here rather than modelled,
and is preserved in this design so that a future batch with a synthetic
malformed-SST corpus and a range-source selector does not have to re-derive any
of it.

## Correctness evidence

Nothing was implemented, so there is no before-and-after to prove. What is
evidenced is that the tree this design is written against is clean and green, and
that the two scaffolds isolate what they claim.

- The two scaffold patches are retained
  ([`scripts/scaffold-nostore.patch`](results/change-0608/scripts/scaffold-nostore.patch),
  [`scripts/scaffold-nowalk.patch`](results/change-0608/scripts/scaffold-nowalk.patch))
  and were applied, built, measured and **reverted**; `git status` was verified
  clean before the gates and the commit.
- The scaffolds are proved to isolate the walk by the self-Ir tables: every
  symbol outside `scan_shared_string_records`'s subtree is identical across the
  three legs, and `MeasuredText::consume`, `walk_one_shared_string` and
  `walk_formatting_runs` are **exactly zero** in `nowalk` on all three fixtures.
- The census probe is cross-checked against production twice: it names the four
  header-refused fixtures change 0576 refuses, and it predicts `isst` 3 for the
  flagship's harness cell, which the scaffold legs report as
  `Invalid SST index: 3 (max: 0)` and the base leg resolves to `string:4:Date`.
- **Gates**, all run in this worktree on the unmodified base tree, tails in
  [`gates.txt`](results/change-0608/gates.txt): `cargo fmt --all --check` clean;
  `cargo clippy -p litchi-xls --all-targets` clean (workspace lints are deny);
  `cargo test -p litchi-xls` 1,382 passed, 0 failed, 1 ignored (the pre-existing
  `#[ignore]` doctest in `writer/core/codec/worksheet.rs`); `cargo doc -p
  litchi-xls --no-deps` clean.

## Validation preserved

Every validation is preserved trivially, because no production code changed. What
the design would do to each is set out in D5 and D6: the SST header checks and
`max_sst_entries` stay at open; the per-string refusals move to first resolve
with byte-identical messages; `max_sst_entries`, `max_global_bytes`,
`max_worksheet_scan_records`, `max_worksheet_scan_bytes`, `max_text_bytes` and
`max_text_cells` are untouched; no ceiling is relaxed; no `unsafe` appears in
either the design or the scaffolds (`crates/litchi-xls/src/lib.rs` keeps
`#![forbid(unsafe_code)]`); the fallible-allocation discipline is kept by
`try_reserve` on every growth; ADR 0003's transaction boundary and ADR 0006's
preservation clause are not engaged, because nothing here is on a write path.

## Limitations

- **Three fixtures, one host, one build, one CPU.** The corpus census covers all
  126 `.xls` and `.xlt` fixtures; every timing and counter figure covers three.
- **Every XLS source in this harness is in memory.** `xls_source_attribution`
  has no `from_path`-observation mode and no full-text or all-cells selector
  (change 0587's standing measurement blocker), so the design's *cost* — D4's
  second read pass on the text path — is modelled from byte counts, never
  measured. That is the weakest part of the evidence here.
- **The two scaffolds are not candidates.** They break resolution by
  construction: `nowalk` and `nostore` both return an empty `entries`, and the
  flagship's `one-cell` on those legs reports `Invalid SST index`. They bound a
  design; they do not implement one, and nothing about their correctness is
  claimed.
- **`base − nowalk` is a ceiling, not a delta.** It removes the walk without
  paying for anything a real deferral would pay for: the lock, the memoised
  refusal, the re-read, the growth reservations. The realisable saving is
  strictly smaller and is not measured, because nothing was implemented.
- **Instruction counts rank work, not latency** (change 0579: 1.24% of
  instructions, 6.19% of cycles). Callgrind counts `rep movsb` per byte, which
  is why the `nostore` leg's memcpy artifact is named above rather than folded
  into its saving.
- **Host quiescence is not established**; both floors are measured, not assumed.
- **The `one-cell` denominators are already superseded.** This record's base is
  `818e58bee`; change 0605 (`perf(xls): walk a whole XLS worksheet in one
  validated scan`) landed on the branch afterwards and rewrites `WorksheetScan`,
  which is 2,806,836 of the 18,603,947 Ir this record reports for `54016.xls`
  `one-cell`. The SST figures and every `open` figure are unaffected — 0605
  touches `workbook/source.rs`'s worksheet scan, not `records.rs` — but the
  2.91% / 17.84% / 10.25% `one-cell` shares are shares of an operation that has
  since become cheaper, and are therefore lower bounds on what the walk is worth
  there now.
- The census counts `LabelSst` records. `RString` (0x00D6) and other legacy
  string records are not counted, and neither is any shared-string reference
  outside a cell — the design's reach could only be *smaller* than stated if one
  existed, never larger, because `idx_max` is already the last entry everywhere.
- The census probe is arithmetic over the container, not `litchi-cfb`: it does
  not apply the FAT validation that refuses `1900DateWindowing.xls` and
  `1904DateWindowing.xls`, which is why it counts 123 and 119 where change 0576
  counts 121 and 117. It also does not decrypt, so the three encrypted fixtures
  are counted as header refusals on their ciphertext count words, which is what
  production does with them too.
- **Not claimed:** any speedup, any regression, any range-source, cold-cache,
  peak-RSS, allocation-count, concurrency-scaling, real-producer or
  cross-platform result; any statement about XLS full text or all cells beyond
  the census's counts; any conclusion about XLS-3, XLS-6, XLS-9, XLS-10 or the
  other items change 0587 ranks near this one.

**Noticed on the way, reported and not acted on.** `SharedStringEntryLocation`
is 16 bytes and the entries are contiguous by construction (D1), so the whole
index is derivable from a `Vec<u32>` of starts plus one final end — a 4× cut in
retained weight, 126,288 B to 31,576 B on `54016.xls`, with `requested` and every
refusal message unchanged, and safe under the 128 MiB default `max_global_bytes`
but needing a fallback if that limit is raised past `u32::MAX`. It moves no
refusal and defers nothing. It is a different item from XLS-2, it was not
measured, and it belongs in a batch with XLS-10 rather than here.

## Retained evidence

[`results/change-0608/README.md`](results/change-0608/README.md) — the census
probe and its full output, the two scaffold patches, the four capture scripts,
every callgrind and `perf stat` output this record cites, the counter table, the
paired wall-clock rounds, both floors, the gate tails, `decision.json` and
`log-sections.md`.
