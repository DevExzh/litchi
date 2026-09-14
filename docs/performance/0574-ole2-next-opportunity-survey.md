# 0574: a source-backed XLS open is now CPU, and most of it decodes strings nobody asked for

Status: attribution only. No production change and `performance_claim: none`.
It re-measures the profile shape changes [0547](changes/0547-ole2-collector-attribution.md)
and [0548](changes/0548-ole2-checkpoint-collector-rejected.md) left, and reports
that changes [0565](0565-xls-globals-single-pass.md),
[0568](0568-xls-worksheet-window.md) and [0570](0570-cfb-fat-run-batching.md)
have moved the OLE2 read path's bottleneck off I/O entirely.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## The one-sentence result

After 0565, 0568 and 0570, a source-backed XLS open spends **75.9% of its wall
time outside the source** and its single largest category is the shared-string
scan, which decodes every shared string in the workbook into a `String` and
throws it away — inclusively **88.33%** of one open of `WithCustomViews.xls`,
**80.81%** of `54016.xls`, and 34.3% of aggregate open time across 117 fixtures.

## Evidence provenance

Every figure here is bound to commit `163ac1bd67f2a0d72c27bfec60e0a2620768cc9d`
and to one `xls_source_attribution` binary,
sha256 `a2955ff9ce8ef1246ada7d8414e72fc299aaf2cbad5db2e6cd6ff85d4ab34449`, built
from that tree at 11:53 local. `crates/litchi-xls` and `crates/litchi-cfb` were
unmodified at that commit and at build time.

A candidate implementing opportunity 1 — a measure-only character walk with a
surrogate-pair validator that keeps the `from_utf16` refusal — appeared in the
working tree at 12:03, after every capture in this record was taken. **This
record is its before leg, not a measurement of it**, and none of these numbers
describe it.

## What was measured

Three scenarios — open and identify, open and list worksheets, open and read A1
of sheet 0 — over three source implementations, on
`test-data/ole/xls/ConditionalFormattingSamples.xls` (1,402,368 bytes), pinned to
CPU 17 with ASLR disabled, 20 warmups and 100 samples. Every logical counter was
identical across all 100 samples of every cell; a varying counter would have
failed the summarizer.

| mode / operation | reads | bytes | `version()` | p50 ns | in-source read ns | in-source `version` ns | outside the source |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `owned-readat` / open | 53 | 565,201 | 29 | 91,840 | 8,950 | 750 | **82,140 (89.4%)** |
| `owned-readat` / list | 53 | 565,201 | 29 | 92,820 | 8,920 | 745 | 83,156 (89.6%) |
| `owned-readat` / one-cell | 61 | 603,115 | 42 | 101,235 | 9,590 | 1,075 | 90,570 (89.5%) |
| `file-source` / open | 53 | 565,201 | 29 | 110,360 | 20,730 | 5,815 | **83,814 (75.9%)** |
| `file-source` / list | 53 | 565,201 | 29 | 109,696 | 20,706 | 5,810 | 83,180 (75.8%) |
| `file-source` / one-cell | 61 | 603,115 | 42 | 124,916 | 23,125 | 8,406 | 93,386 (74.8%) |
| `tracked-file` / open | 53 | 565,201 | 29 | 114,740 | 20,290 | 750 | 93,700 (81.7%) |
| `tracked-file` / one-cell | 61 | 603,115 | 42 | 130,901 | 22,990 | 1,100 | 106,810 (81.6%) |

The read and byte counts reproduce change 0565's retained figures exactly, which
is the control: 0568 and 0570 move nothing on this fixture, as both records
predicted.

**Open and list are identical read for read and byte for byte.** Listing
worksheets costs nothing beyond the open, because every worksheet name is already
resident. **One cell costs 8 reads and 37,914 bytes more than the open** — the
0568 window on the selected sheet's own validated region — so the worksheet scan
is bounded and is no longer where the work is.

Hardware counters, isolated by differencing an 1100-sample and a 100-sample child
and dividing by 1000, the same way change 0564 isolated syscalls:

| mode / operation | instructions/op | branches/op | branch misses/op |
| --- | ---: | ---: | ---: |
| `owned-readat` / open | 1,645,028 | 328,644 | 1,019 |
| `owned-readat` / one-cell | 1,832,731 | 365,627 | 1,031 |
| `file-source` / open | 1,818,895 | 362,098 | 1,217 |
| `file-source` / one-cell | 2,041,457 | 405,102 | 1,273 |

1.65 million instructions to return sixteen sheet names.

## Where the instructions go

Callgrind, same isolation method, on three fixtures chosen for different globals
composition. Shares are of the per-open instruction delta.

Self cost, summed over the symbols in each category:

| category | `ConditionalFormattingSamples.xls` | `WithCustomViews.xls` | `54016.xls` |
| --- | ---: | ---: | ---: |
| **shared-string scan** | **10.68%** | **82.37%** | **71.58%** |
| `memset` + `memcpy` | 44.20% | 6.12% | 8.51% |
| allocator | 10.28% | 6.15% | 8.56% |
| formatting parse | 7.25% | 1.43% | 4.62% |
| CFB whole-container validation | 5.64% | 0.34% | 0.60% |
| FAT chain prefix walk | 4.96% | 0.20% | 0.26% |
| Ir per open | 2,801,797 | 5,601,658 | 18,426,423 |

Taken *inclusively*, so that the allocator traffic the scan causes is counted
where it is caused, `scan_shared_string_records` is **382,422 Ir (13.65%)**,
**4,947,696 Ir (88.33%)** and **14,890,837 Ir (80.81%)** of those three opens. Both
framings are reported; the inclusive one is the size of the opportunity and the
self one is where the instructions are executed.

Top symbols on the two string-heavy fixtures:

| symbol | `WithCustomViews.xls` | `54016.xls` |
| --- | ---: | ---: |
| `String::from_utf16` | **47.05%** | **31.38%** |
| `SstCursor::read_characters` | 32.22% | 23.90% |
| `scan_shared_string_records` | 0.97% | 6.58% |
| `parse_one_shared_string` | 1.01% | 5.06% |

Per-open call counts, from the same isolation pair's caller tree:

| edge | calls per open | Ir per open |
| --- | ---: | ---: |
| `GlobalsBuffer::ensure` → `__memset_avx2_unaligned_erms` | 20 | 548,308 |
| `read_stream_range` → `next_chain_sector` | **5,796** | 139,104 |
| `SstCursor::read_characters` → `String::from_utf16` | **303** | 135,609 |

`GlobalsBuffer::ensure` is the caller of **98.4%** of every `memset` instruction in
the flagship profile: 548,308 of 556,908. The 5,796 chain steps are the cost of
re-walking the allocation chain from its first sector on each of the 20 fills. The
303 `from_utf16` calls are 303 shared strings decoded and dropped.

## The mechanism, read from the source

`scan_shared_string_records` (`litchi-xls/src/records.rs:936`) builds an index of
`(start, end)` logical offsets so that a later cell lookup can fetch one entry.
It obtains each boundary by *fully parsing the string and discarding it*
(`records.rs:1023-1028`):

```rust
for string_index in 0..unique_count {
    let start = cursor.logical_position();
    parse_one_shared_string(&mut cursor, string_index)?;
    let end = cursor.logical_position();
    entries.push(SharedStringEntryLocation { start, end });
}
```

`parse_one_shared_string` reaches `read_characters` (`records.rs:619`), which
allocates a `Vec<u16>` per string, pushes one code unit at a time, and finishes
with `String::from_utf16(&characters)` (`records.rs:669`) — a second allocation
and a full UTF-16 to UTF-8 transcode. The returned `String` is dropped
immediately. Per shared string that is two heap allocations, two frees, a
per-character loop and a transcode, to advance a cursor.

`SharedStringSstScan` itself retains only offsets, and the scan's own output is
correct and small. The waste is entirely in how the offsets are obtained.

## Ranked opportunities

### 1. Stop materializing shared strings during the open-time SST scan

**Mechanism.** Replace the discard-the-result call with a measure-only walk that
advances the cursor through the same framing — character count, flags, rich-text
run count, phonetic block, `Continue` boundaries and continuation flags — without
building a `Vec<u16>` or a `String`. `records.rs:619-670` already computes every
length it needs; it also materializes.

**Expected saving, measured.** 88.33% of one open of `WithCustomViews.xls`
(≈204 µs of its 231 µs p50), 80.81% of `54016.xls` (≈572 µs of 708 µs), 13.65% of
the flagship. Corpus-wide the regression below attributes **34.3%** of aggregate
open time across 117 fixtures to SST bytes, at 2,541.7 ns per KiB of SST. A
measure-only walk does not recover all of it — the framing walk remains — but
`String::from_utf16` alone is 47.05% and 31.38% of those two opens, and it is
pure waste.

**ADR constraints.** ADR 0005 states that opening performs "container,
relationship/catalog, security, and mandatory structural validation" and that
"semantic payloads load lazily". A shared-string table is a semantic payload, so
both a measure-only walk and full deferral are *aligned* with ADR 0005 rather
than in tension with it. No accepted ADR requires eager SST decoding.

**What it changes, and why it is not landed here.** `String::from_utf16` returns
`Error::Encoding("UTF-16 decoding error: …")` on a lone surrogate, today at open.
A measure-only walk that does not validate UTF-16 well-formedness moves that
refusal to first shared-string access; a walk that does validate keeps the
refusal but gives back less. That is a deliberate, testable contract change of
exactly the kind changes 0565 and 0566 froze in a design record before
implementing. This record does not establish it.

**Falsified if.** A measure-only walk cannot reproduce byte-identical `entries`
over every fixture with an SST — 123 of 126 carry one, three of which are
encrypted and do not open (`model/sst-counts.json`); or the saving on the flagship
is inside noise because its SST is only 4,603 bytes, which is why the ranking
rests on `WithCustomViews.xls` and `54016.xls`, not on it.

**Concentration, stated plainly.** The SST is 1.5% of globals for the median
opened fixture and over a quarter for only 7 of 117. What makes this the top
opportunity is not how often it is large but how expensive its bytes are — 2,542
against 53 ns/KiB — which is why it is 34.3% of aggregate open time while the
opaque over-read, **seven times** its byte volume across the same fixtures
(2,970,299 against 423,448 bytes), is 5.1%.

### 2. Stop reading globals payloads the semantic pass never interprets

**Mechanism.** `parse_globals` (`workbook/source.rs:1659`) reads `[0, global_end)`
in full. Its semantic pass consumes exactly twelve record kinds: `BOF`, `EOF`,
`FilePass`, `CodePage`, `BoundSheet8`, `SST` with its `Continue` run, and — through
`Formatting::parse_globals` (`number_format/codec.rs:57`) — `Date1904`, `Format`,
`XF`, `XFCRC`, `XFExt` and `DXF`. Every other payload is buffered, framed twice
and dropped. Skipping them uses the same cursor `skip_forward` primitive the 0568
worksheet scan already relies on.

**Expected saving, measured and modelled.** A static model over 126 XLS fixtures:

| | bytes | share |
| --- | ---: | ---: |
| globals bytes read today | 3,999,694 | 100% |
| four-byte record headers | 106,428 | 2.66% |
| consumed record payloads | 890,268 | 22.26% |
| **touched closure** | **996,696** | **24.92%** |
| **payloads never interpreted** | **3,002,998** | **75.08%** |

On the flagship this is 530,960 of 551,377 globals bytes, **95.2% of them a
`MsoDrawingGroup` record and its `Continue` chain** — 16,456 bytes of record and
508,383 bytes of continuation, verified to continue that record and no other. At
the regression's measured 53.4 ns per KiB of skipped payload that is **27.7 µs of
the flagship's 91.8 µs open, 30.2%**, and 155 µs across the 117-fixture aggregate,
**5.1%**.

**This is a concentrated result and the aggregate share hides it.** The median
fixture consumes **66.6%** of its globals (p10 42.5%, p90 79.0%), and 103 of 126
consume more than half. **90% of all skippable bytes live in 15 of the 126
fixtures.** The corpus is the right one for asking how often this helps and the
wrong one for asking how much — the same shape change 0570 recorded for FAT
batching.

**The cost is reads.** A modelled schedule that merges needed spans separated by
under one 512-byte sector takes 65 logical range requests on the flagship against
today's 19 fills; the physical split of those requests into contiguous runs is not
modelled, so the real read count is higher on both sides. At the ~116 ns per
syscall change 0564 derived, the flagship trade is about +5.3 µs of syscall
against −27.7 µs of bytes. It is the inverse of change 0568's trade and must be
stated that way.

**ADR constraints.** The "buffered but never published" contract change 0565
established already covers bytes resident behind a late `FILEPASS`; this extends
it to bytes never fetched. `max_global_bytes` bounds retained globals and would
bind more tightly, not less. ADR 0006's preservation contract is untouched: this
is the source-backed read path, which publishes no bytes.

**Falsified if.** The read-count rise costs more than the byte saving on a
fixture whose globals are dense. That is the majority case, not the exception:
103 of 126 fixtures consume more than half their globals and would pay the extra
reads for little. A gate on measured density — the primitive change 0568 already
built for the worksheet scan — is a precondition, not a refinement.

### 3. Give the globals scan a retained stream cursor

**Mechanism.** `SharedOleFile::read_stream_range` (`litchi-cfb/src/shared.rs:801`)
walks the allocation chain from the stream's **first** sector on every call:

```rust
for _ in 0..first_ordinal {
    sector = next_chain_sector(&self.index.fat, sector, "FAT")?;
```

(`shared.rs:1933`.) The globals scan issues 20 such calls at increasing offsets,
so the prefix walk is quadratic in the number of fills. Change 0566 identified
exactly this and gave the worksheet scan a `SharedOleStreamCursor` for it; the
globals scan still uses the range reader.

**Expected saving, measured.** 5,796 chain steps per flagship open where the
bytes require about 1,077. `next_chain_sector` is 4.96% and `read_stream_range`
a further 4.30% of flagship open instructions — 139,104 and 120,463 Ir. The
recoverable part is the prefix walk, so up to about 4.96%, roughly 4.5 µs. On
fixtures with small globals it is 0.20% to 0.26% and worth nothing.

**ADR constraints.** None new. The cursor is an existing CFB primitive already
used on this path's sibling.

**Falsified if.** Read counts, read bytes or source-version observations change —
they must not, and the counted evidence in this record is the before leg. Also
falsified if the mini-stream-resident case (four fixtures) behaves differently
through a cursor than through the range reader.

### 4 to 8, recorded and not recommended now

| # | Opportunity | Measured size | Why not now |
| --- | --- | ---: | --- |
| 4 | Stop zero-filling the globals buffer. `self.bytes.resize(finish, 0)` (`source.rs:1620`) memsets every byte the very next line overwrites; `GlobalsBuffer::ensure` causes 98.4% of the profile's `memset`. | up to 19.88% of flagship open Ir | Needs a read API over `&mut [MaybeUninit<u8>]`; avoiding it otherwise needs `unsafe`, which GOAL rule 10 forbids without an ADR-permitted owner. Opportunity 2 subsumes most of it by shrinking the buffer 27-fold. |
| 5 | Frame the globals once. The scan loop frames all 621 records, then `BiffRecords::with_limits(&bytes, …)` (`source.rs:1809`) frames all 621 again. | 1.51% of flagship open Ir | Small, and entangled with opportunity 2, which changes the buffer's shape. Do it inside that change, not before it. |
| 6 | Make CFB whole-container validation proportional to the touched closure. `validate_stream_allocations` (`file.rs:1110`) walks every stream's full chain at open. | 5.64% / 0.34% / 0.60% of open Ir | ADR 0005 names mandatory structural validation as an open-time obligation and GOAL rule 12 forbids weakening ownership and overlap defenses. Not available. |
| 7 | Coalesce stream reads across non-contiguous sector runs. | zero on this corpus | Reads are already at the physical run minimum; coalescing would read sectors owned by other streams, which is the ownership boundary ADR 0006 protects. |
| 8 | The DIFAT loop (`file.rs:846`). | zero | See below. |

## The explicit checks

**Does a one-cell read still do work proportional to the whole workbook stream?**
**No for the stream, yes for the globals.** One cell reads 603,115 of the
1,314,225-byte `Workbook` stream, 45.9%. The worksheet component is 8 reads and
37,914 bytes, bounded by the selected sheet's own validated region exactly as
change 0568 built it. But the open underneath it reads the whole 551,377-byte
globals substream, 42% of the stream, of which 3.70% is the touched closure. The
0568 window holds; the globals pass is where the whole-document proportionality
now lives.

**Is the MiniFAT or ministream materialized in full for files that touch only a
few small streams?** **The ministream is not; the MiniFAT table is, and it does
not matter here.** Measured over the four fixtures whose `Workbook` stream lives
in the CFB mini stream — three distinct files, since `SimpleWithColours copy.xls`
is byte-identical to its original — subtracting the exact CFB structural bytes
from change 0565's retained geometry:

| fixture | reads | payload bytes | globals end | ministream size |
| --- | ---: | ---: | ---: | ---: |
| `SimpleWithColours.xls` | 12 | 3,177 | 1,902 | 3,776 |
| `WithCheckBoxes.xls` | 17 | **2,465** | **2,465** | 12,224 |
| `WithExtendedStyles.xls` | 11 | 1,576 | 1,386 | 3,456 |

Payload bytes track the globals substream, not the root mini stream, on every
fixture; `WithCheckBoxes.xls` reads exactly its globals and one fifth of its
ministream. `SharedOleFile::read_stream_range` takes the direct MiniFAT path and
its doc comment's claim that it "never materializes … the root mini-stream" is
confirmed by measurement. The MiniFAT *table* is loaded in full and
unconditionally at open (`file.rs:764-766`), but no fixture in `test-data/ole`
has more than two MiniFAT sectors, so it is at most 1,024 bytes and one or two
reads.

**Are directory and stream reads issuing more positional reads than the sector
runs require?** **No, and neither batches across a gap.** `read_sectors_batched`
(`file.rs:2145`) breaks a run at the first non-successor (`file.rs:2160-2168`) and
issues one read per run; `read_sector_run` (`shared.rs:2344`) reads one
already-identified run, with discovery in `read_chain_into` (`shared.rs:2269`)
ending a run on `next != sector + 1`; `read_stream_range`'s partial path has the
same inline loop. All three are contiguous-only. The flagship's 53 reads for
565,201 bytes is within one read of the physical-span minimum its geometry allows.
Gap-spanning is opportunity 7 and is not recommended.

**Is the DIFAT loop ever hot on a real fixture?** **No, and change 0570's claim is
verified.** Zero of the 98 container fixtures under `test-data/ole` declare any
DIFAT sector; the largest FAT is 23 sectors (`ole/doc/picture.doc`, 1,448,448
bytes). The header's 109 DIFAT entries map 13,952 FAT sectors at a 512-byte
sector size, so no file under about **7.14 MB** can need one, and none under about
**457 MB** at 4,096 bytes. The loop is not merely cold on this corpus, it is
unreachable for the entire size class the corpus occupies, and the
`expected_fat_sectors` mismatch check past it remains untested.

**Is change 0547's collector shape still the profile?** **Not in the same
proportion, and the scope must be stated to compare at all.** Measured on the
flagship, inclusive, per open:

| scope | Ir per open | share |
| --- | ---: | --- |
| whole source-backed open | 2,801,797 | 100% |
| XLS source constructor (`from_shared_ole_file_with_limits`) | 2,475,302 | 88.35% of open |
| CFB constructor (`OleFile::open_with_limits`) | 273,821 | 9.77% of open |
| exact-chain collector (`SectorChainScratch::collect_exact`) | 105,475 | **38.52% of the CFB constructor**, 4.26% of the XLS constructor |

Applying change 0547's own measured ratio — the visited category is 41.15% to
41.17% of collector self Ir — to today's 93,714 Ir of collector self work puts the
visited-map per-step checks at roughly 38,500 Ir per open, about **1.4% of the
open** and about 14% of the CFB constructor, against 0547's 20.37% of the
constructor it measured. The collector is still the largest single component
*inside* the CFB constructor, but the CFB constructor is now under a tenth of the
open. **This is a change of proportion, not a refutation of 0547**: its absolute
magnitudes — 11,316,345 constructor and 5,601,140 collector self Ir across five
timed dumps, which is 2,263,269 and 1,120,228 per dump *if* a dump is one open,
and that is an inference this record cannot check — cannot be bound to this tree,
because its fixture and harness binding is not reproducible from the retained
record, and 0547's rejected successor 0548 is correctly absent from the tree. The
conclusion for planning is the one that matters: a second
attempt at the collector would be competing for about 1.4% where opportunity 1 is
competing for 34.3%.

## The corpus regression

One open of every XLS fixture under `test-data` that opens — 117 of 126 — against
the static composition model. Fitting p50 open time to the three byte classes:

```
p50_open_ns = 4,323 + 0.0522*skipped + 2.4821*sst + 2.6311*other_closure
R^2 = 0.9933
```

| byte class | cost | aggregate | share of 3,065,642 ns |
| --- | ---: | ---: | ---: |
| opaque globals payload, read and discarded | 53.4 ns/KiB | 154,942 ns | 5.1% |
| SST bytes | 2,541.7 ns/KiB | 1,051,055 ns | **34.3%** |
| other consumed closure | 2,694.2 ns/KiB | 1,353,858 ns | 44.2% |
| fixed per open | 4,323 ns | 505,786 ns | 16.5% |

**A consumed globals byte costs 48 to 50 times what a skipped one costs.** That is the
finding that sets the ranking, and it is the opposite of what the read counts
alone suggest. The regression cannot fully separate SST bytes from the rest of
the closure — the two coefficients are within 6% of each other and the classes are
correlated with globals size — so the per-fixture callgrind attributions, not this
regression, are the evidence for opportunity 1's size. The regression's role is
to establish that consumed bytes, not read bytes, drive open cost.

## No code changed

No file under `crates/` was modified. Opportunity 1 changes when a malformed
UTF-16 shared string is refused; opportunity 2 changes which globals bytes are
read; opportunity 3 changes which CFB primitive the globals scan uses. None is
small enough and obviously safe enough to land as a side effect of a survey, and
each deserves the frozen-design treatment changes 0566 and 0565 gave their
predecessors.

## Limitations

No before/after pair, no A/B/B/A latency matrix, no host-quiescence log, and
therefore no admissible speedup claim. The retained latency figures are
single-leg attribution medians used only to size opportunities against one
another; they are warm-cache, page-cached, on a staged immutable copy, on one
host, from one binary, and they are not comparable across the three source
implementations for anything but their own documented scope.

No cold-cache, physical-device, remote or range-source, peak-RSS,
allocation-profile, concurrency-scaling, real-producer or cross-platform result
is claimed. No DOC or PPT scenario was measured at all: the three scenarios are
XLS, and the CFB findings (opportunity 3, the DIFAT verification, the batching
answer) are the only part that transfers to them without fresh measurement.

Callgrind instruction counts for `__memset_avx2_unaligned_erms` and
`__memcpy_avx_unaligned_erms` are **upper bounds**, because Valgrind instruments
the ERMS string loops per iteration where the hardware retires far fewer
instructions. That is why opportunity 4 is sized from those counts only as a
ceiling and opportunity 2 is sized from the regression's 53.4 ns/KiB instead.
Callgrind's 2,801,797 Ir per flagship open against `perf`'s 1,645,028 is that
same effect and is reported rather than reconciled.

The static closure model reads the twelve consumed record kinds out of two match
statements. If a future record kind is consumed and the model is not updated, its
75.08% skippable share is an overestimate. It is validated against the measured
globals end and record counts on every fixture, not against an instrumented run.

Nine of 126 fixtures do not open, including the three encrypted ones, and are
excluded from the regression rather than imputed.

[Captures, models, profiles and the replay script](results/change-0574/README.md).
