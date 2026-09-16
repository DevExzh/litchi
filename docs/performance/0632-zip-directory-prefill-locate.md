# 0632: the central directory is read once, and the buffer it lands in is sized to it — every OOXML open loses a request and a 64 KiB scratch

Status: retained, implemented in `soapberry-zip` (three production files, no
public API added or changed). `performance_claim: none` — no claim-registry
entry in this wave; the paired medians and the deterministic counts below are
reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **ZIP-6** of
[0587](0587-remaining-opportunity-survey.md) — the bounded tail-window locate —
re-priced by change [0623](0623-zip-structural-span-accessor-and-prefetch.md),
together with the first half of **ZIP-8**, the 64 KiB central-directory scratch
that is zero-filled although the directory size is known.

**The baseline moved twice before this landed.** Change
[0611](0611-zip-single-read-per-member.md) took the 132-member workbook's
source-backed open from 89 requests to 45, and change 0623 took it to **10**.
Three of those 10 are the locate, so everything here is measured and stated
against 10.

## What was changed

One crate's production code, three files, 244 inserted lines and 39 deleted.

**The two reads that begin at the same offset become one.** After change 0623 a
source-backed open of `ConditionalFormattingSamples.xlsx` issues these three
structural requests:

| # | site | offset | bytes |
| --- | --- | --- | ---: |
| 1 | `ZipLocator::locate_in_reader`'s fixed-EOCD probe | `len - 22` | 22 |
| 2 | `ZipLocator::finish_locate_in_reader`'s first-central-record probe | `central_dir_offset` | 46 |
| 3 | `ZipEntries::next_entry`'s first refill, from `IndexedArchive` | `central_dir_offset` | 9,724 |

Requests 2 and 3 begin at the same offset: the 46 bytes of request 2 are the
first 46 bytes of request 3. One `ReadAt` call now fetches that range.

**`crates/soapberry-zip/src/locator.rs`.** `ZipLocator` gains a `pub(crate)`
`directory_prefill(bytes)` option, off by default, and a `pub(crate)`
`locate_in_reader_prefilling_directory`. When the option is set,
`finish_locate_in_reader`'s first-central-record probe reads
`min(central_directory_size, W)` bytes at the offset it was going to probe,
into a fallibly reserved `DirectoryPrefill`, and hands that buffer back. The
probe itself is unchanged: `ZipFileHeaderFixed::parse` of the same 46 bytes at
the same offset decides the same way, and the base-offset fallback at
`head_eocd_offset - central_dir_size` runs under exactly the same condition, as
the same kind of read. `locate_in_reader` keeps its signature, its behaviour and
its 46-byte stack probe — it is now one line that discards the prefill, and with
the option off no prefill is ever taken.

**`crates/soapberry-zip/src/archive.rs`.** `ZipEntries` gains an explicit
`spill_threshold` and a `pub(crate)` constructor, `entries_for_index`, that
starts the scan from a prefilled buffer. Every existing constructor sets
`spill_threshold = buffer.len()`, which is the boundary this iterator has always
used, so no existing caller moves. `entries_for_index` pins it at
`RECOMMENDED_BUFFER_SIZE` — see *Why it is sound*, point 4.

**`crates/soapberry-zip/src/office.rs`.**
`IndexedArchive::from_reader_with_limits_and_policy` asks for a prefill of
`W = RECOMMENDED_BUFFER_SIZE` (64 KiB), releases the locator scratch as soon as
the locate returns, and uses the prefill buffer as the central-directory scan
buffer. The **second** 64 KiB `try_reserve_exact` + `resize(.., 0)` is gone. The
public `from_zip_archive_with_limits_and_policy`, which takes an archive the
caller located itself, takes no prefill but does size its scan buffer to
`min(central_directory_size, 64 KiB)`.

Nothing else changed. No new `unsafe`, no new dependency, no public API added,
removed or altered, no weakened limit or defence, and no change to any writer,
publisher, strict-layout or preservation path.

**Three of change 0623's assertions are updated.**
`crates/litchi-opc/tests/structural_prefetch.rs` pins the *absolute* request
count of three opens at 4, 10 and 14 reads, and three of each count are the
locate. Each falls by exactly one, to 3, 9 and 13, and the comments that name
"three locator reads" now name two. No other line of those tests changed, and
nothing else in the repository asserts an absolute ZIP open request count.

### Why `W` is 64 KiB

`RECOMMENDED_BUFFER_SIZE` is the size of the buffer the scan uses today, so
choosing it makes the prefilled scan state-identical to today's scan after its
own first read (point 3 below), and it is the per-request ceiling change 0623
already set for one speculative read. Change
[0573](0573-zip-single-local-header-read.md) censused the same 533 containers
member by member — 14,744 members, largest local variable region 539 bytes — to
set its 640-byte local window; this record censuses their *tails*, which is the
half 0573 did not need (`results/change-0632/census/`), and the window covers
the whole directory on every one of them:

| central directory, 533 containers | |
| --- | ---: |
| minimum | 287 B |
| median | 1,217 B |
| p90 / p99 | 4,904 B / 7,166 B |
| **maximum** | **9,724 B** (`ConditionalFormattingSamples.xlsx`) |
| containers at or below 16 KiB | **533 of 533** |
| containers with a non-empty archive comment | **0** |
| containers with ZIP64 end-of-central-directory metadata | **0** |
| containers whose first central record is not at the declared offset | **0** |
| containers whose EOCD is not at `len - 22` | **1** (`bug62513.pptx`, one trailing byte) |

A directory larger than `W` is not a correctness case: the window covers its
head, the scan reads the rest, and the request count is what it is today. That
is measured by a test, not argued.

## Why it is sound

Let `n = min(central_directory_size, W)` and `d` the offset being probed.

**1. The probe sees the same bytes, and fails on the same conditions.** Today:
`read_exact_at(&mut [0u8; 46], d)`, then parse. After:
`try_read_at_least_at(&mut buf[..n], n, d)`, then parse `buf[..46]`. Both stop
on exactly two conditions — an `Err` from the source, or a zero-length read
before the bytes are in hand — so a probe that fails today fails after, and one
that succeeds today parses the same 46 bytes. A prefill that returns fewer than
46 bytes is discarded and the probe is refused, which is what `read_exact_at` of
46 bytes does on the same source. A declared directory shorter than one fixed
central record takes **no** prefill and keeps the 46-byte stack probe.

**2. The base-offset fallback is unchanged.** When the probe at the declared
offset does not parse, the same second probe runs at
`head_eocd_offset - central_dir_size` and updates `base_offset` and
`central_dir_offset` under exactly today's condition. A prefill whose start is
not the archive's final `directory_offset()` is discarded by the index, so a
re-aimed archive can never be scanned from a buffer read at the wrong place.

**3. The scan starts in the state its own first read would have left.** Today
`ZipEntries::next_entry`'s first refill calls
`read_at_least_at(&mut buffer[..max_read], 46, d)` with
`max_read = min(central_dir_size, buffer.len())` and leaves `pos = 0`,
`end = read`, `offset = d + read`. With a prefill of the same width the iterator
starts at `pos = 0`, `end = valid`, `offset = d + valid` — the same numbers, from
the same bytes. Every later refill, every metadata charge, every parse and every
error is then reached from the same buffer contents at the same logical
position. This is why the differential below finds the two reports byte-identical
rather than merely equivalent.

**4. The oversized-record spill boundary is pinned, not moved.** `next_entry`
reads a record whose variable fields exceed the caller's buffer into an owned
spill buffer, so that boundary is a function of the buffer size. Sizing the scan
buffer to the directory would move it: a record declaring
`central_dir_size < variable_length <= 65,536` refuses with `BufferTooSmall`
today, through the refill, and would refuse with `Eof` after, through the spill.
`entries_for_index` therefore pins `spill_threshold` at `RECOMMENDED_BUFFER_SIZE`
instead of tracking the buffer. With it pinned, a record larger than 64 KiB
spills exactly as before, and anything smaller that does not fit refuses with
`BufferTooSmall` exactly as before — `an_oversized_variable_field_keeps_its_typed_refusal`
compares the two directly and the differential confirms it over 22,875 inputs.

**5. Nothing is trusted because it came from the buffer.** The prefill is a cache
of source bytes. Every central record is parsed, every metadata byte charged,
every name-length and metadata ceiling checked, and every physical bound applied
by exactly the code that does so when the scan reads those bytes itself, in the
same order.

**6. Bounded and fallible.** The prefill is `try_reserve_exact`d, is at most `W`
and at most `central_directory_size`, and is released with the index
construction; a reservation failure yields no prefill rather than an error. The
open's peak scratch **falls**: two 64 KiB buffers become one 64 KiB locator
buffer, released before the scan, plus one buffer the size of the directory.

**7. One movement is accepted, and it is change 0611's.** Because this change is
below `litchi-opc`, a *managed* open also loses the request. It therefore
reserves and charges `Resource::InputBytes` once for `central_directory_size`
where it charged 46 and then `central_directory_size`, and makes one fewer
`ExecutionContext::check` cancellation observation. The merged read is the read
the scan had to issue, one for one, with a strict subset of it removed — which is
exactly the argument change 0611 used and change 0623 could not — and the
movement is in the direction of less work: a finite input budget within 46 bytes
of the boundary now admits an open it refused, and none can refuse one it
admitted. See *Limitations*.

**ADR reading.** ADR 0005's bounded-resource clause is met by a named ceiling, a
fallible reservation and a lifetime that ends with the index; its "cache
behavior is semantically invisible" clause is met by the byte-identical
differential report and by the fact that the cached bytes are the bytes the very
next read was going to fetch. ADR 0006 is untouched: every check keeps its
position and its identity, and no output byte is produced on this path. ADR 0011
keeps its line: this is entirely inside `soapberry-zip`, the ZIP grammar owner,
and nothing new crosses the boundary — `litchi-opc` sees one fewer `read_at`
call and 46 fewer bytes. ADR 0003's "validation must not mutate" is untouched;
`ArchiveValidationPolicy` is not read by any changed line.

**`ReadLimits` is not reached by any changed line.** `litchi-opc`'s read
ceilings live a layer above this change and are charged by the same code in the
same order; the 533-container open differential opens every container with
`ReadLimits::default()` and reports identical verdicts, identical relationship
and part catalogs and identical decoded payload CRCs.

**Which contracts are untouched.** The EOCD search rules keep their `len - 22`
fast-path gate (`comment_len == 0 && !is_zip64()`), their backwards search and
their `max_search_space`; comment handling, the ZIP64 locator and the ZIP64
record resolution are not touched at all. `ZipLocator::locate_in_reader`,
`locate_in_file`, `locate_in_slice`, `ZipArchive::entries`,
`entries_with_metadata_limit`, `PreservationIndex` and every strict-layout path
of [0580](0580-zip-target-scoped-strict-layout.md) and
[0583](0583-zip-local-size-span-bound.md) keep their exact grammar, because the
option is off for all of them. Change 0611's per-member span, change 0623's
structural prefetch and `ZipOperationAccounting`'s counters are unchanged.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0.
Base `c7326f680`, branch `perf/0632-zip-tail-window-locate`. Every measured
process pinned to CPU 8. Seven other agents were building and measuring on the
host throughout, at a load average between 15 and 42; every window carries its
own A/A floor and, for the paired timing, a B/B floor as well.

### Requests and bytes (deterministic counts, measured)

Change 0587's probe as change 0623 left it, re-run verbatim on both legs against
the same three fixtures (`results/change-0632/probe/`, outputs in
`results/change-0632/counts/`). The before leg reproduces change 0623's retained
after-counts exactly.

| scenario | before | after | requests | bytes | open allocation |
| --- | --- | --- | ---: | ---: | ---: |
| open `ConditionalFormattingSamples.xlsx` (132 members, 9,724 B directory) | 10 req / 26,066 B | **9 / 26,020** | **−10.0%** | −46 | 2,312,260 → **2,256,447** (−55,813) |
| open `shapes.pptx` (48 members, 3,684 B directory) | 6 / 11,191 | **5 / 11,145** | **−16.7%** | −46 | 603,186 → **541,333** (−61,853) |
| open `comment.docx` (10 members, 635 B directory) | 6 / 1,690 | **5 / 1,644** | **−16.7%** | −46 | 430,955 → **366,053** (−64,902) |
| first read of `/xl/worksheets/sheet1.xml` | 1 / 1,305 | 1 / 1,305 | 0 | 0 | — |
| first read of `/ppt/slides/slide1.xml` | 1 / 4,807 | 1 / 4,807 | 0 | 0 | — |
| first read of `/word/document.xml` | 1 / 571 | 1 / 571 | 0 | 0 | — |
| all 90 XLSX parts, either order | 90 / 628,668 | 90 / 628,668 | 0 | 0 | — |
| all 27 PPTX parts, either order | 27 / 57,677 | 27 / 57,677 | 0 | 0 | — |
| all 7 DOCX parts, either order | 7 / 3,498 | 7 / 3,498 | 0 | 0 | — |
| `stream_to` of each of the three parts | 21 / 1,875, 16 / 5,227, 12 / 782 | identical | 0 | 0 | — |

Only the open line moves; the `cd-probe(46)` class the probe already names
disappears and nothing else changes. `version()` observations are **unchanged**
on all three fixtures (4 at open, and the same on every later phase) — the
merged read is one source contact where there were two, and the source-version
fence is charged in the same place. The allocation saving is
`RECOMMENDED_BUFFER_SIZE − central_directory_size` in each case, to the byte.

### The open differential over every ZIP container under `test-data`

533 containers — the same census change 0623 and change 0582 take. Both legs
record, per container: the open's verdict, its request, byte and
source-observation cost, every package and part relationship with its id, type,
target and mode, every admitted Part with its content type and relationship
count, every non-part member with its reason, and the decoded length and CRC-32
of every Part. 18,494 lines per report.

**The two reports differ on 533 lines, and every one of them is an
`open-cost requests=` line.** No verdict, no error identity, no part, no
relationship, no non-part member, no decoded payload and no source-observation
count differs anywhere in the corpus.

| | before | after |
| --- | ---: | ---: |
| open requests, 533 containers | 2,569 | **2,036** (−20.7%) |
| open requests, 338 OOXML-extension containers | 1,984 | **1,646** (−17.0%) |
| open bytes, 533 containers | 2,391,391 | **2,366,873** (−24,518) |
| `version()` observations | 2,132 | **2,132** |
| containers costing **more** requests | — | **0** |
| containers costing fewer requests | — | **533** |
| containers reading **more** bytes | — | **0** |
| containers reading fewer bytes | — | **533** |
| per-container request ratio | — | min 0.667, median 0.800, max 0.957 |

The saving is exactly one request and exactly 46 bytes on **every** container:
no container in the corpus reaches the base-offset fallback, has a comment, or
has a directory larger than the window, so no container pays for a prefill it
does not use.

### The read-grammar differential of change 0611, in full

Change 0582's 22,875-input corpus, regenerated byte for byte from `RNG_SEED`
`0x05820580`; change 0611's extended harness, which adds `I.read_entry` and the
slice-backed control `R.read` to change 0582's seven strict-layout APIs; both
limit profiles (`fuzz` and `wide`); both read directions per API
(`order_independent`) and the memo-stability sweep; both builds from `git
archive` trees with their own target directories.

| | |
| --- | ---: |
| inputs | 22,875 |
| member verdicts examined | 2,886,786 |
| report lines per leg | 3,801,502 |
| class A / class B / class E divergences | **0 / 0 / 0** |
| oracle failures, panics | **0 / 0** |
| `cmp report-before.txt report-after.txt` | **byte-identical** (`sha256 358e306f…3894` on both) |

The two reports are not merely equivalent, they are the same file. That is the
falsification criterion the frozen design set, met in the strongest available
form.

### Paired timing on a latency-bearing transport

Both binaries `--release --locked` from the same sources with the same flags and
the same `Cargo.lock`, copied out of their Cargo target directories before the
first leg (change 0627: a concurrent build relinked one mid-run), `taskset -c 8`,
leg order A1 B1 B2 A2 with an A/A floor A3 A4 in the same window; the B/B floor
is B1 against B2, two legs of the same after binary.

**The simulated transport of changes 0493 and 0572** — 1 ms of fixed service per
physical request, 100 MiB/s, 64 KiB maximum physical range — so what is priced is
the request count, and the request count is deterministic.

*(a) The harness's own range-source selectors*, 30 samples and 5 warmups per case
per leg (`results/change-0632/timing/`). Physical requests per timed iteration,
from the simulator's counters, are identical across all samples of every leg:

| case | requests before → after | bytes before → after |
| --- | --- | --- |
| `opc_range_source_open` | 4 → **3** | 987 → **941** |
| `opc_range_source_open_main_read` | 5 → **4** | 1,142 → **1,096** |
| `xlsx_range_source_open` | 6 → **5** | 1,875 → **1,829** |
| `xlsx_range_source_first_cell` | 10 → 10 | 666 → 666 |
| `xls_range_source_open` (OLE2 control) | 16 → 16 | 110,242 → 110,242 |

| case | A1→B1 p50 | A2→B2 p50 | predicted | floor A3→A4 | floor A1→A2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `opc_range_source_open` | **−24.82%** | **−24.91%** | −25.0% | +0.01% | +0.09% |
| `opc_range_source_open_main_read` | **−19.91%** | **−19.90%** | −20.0% | −0.05% | +0.01% |
| `xlsx_range_source_open` | **−16.68%** | **−17.02%** | −16.7% | +0.10% | +0.30% |
| `xlsx_range_source_first_cell` | −0.03% | +0.03% | 0 | +0.03% | +0.02% |
| `xls_range_source_open` (OLE2 control) | +1.52% | −0.09% | 0 | +0.01% | +0.15% |

The two paired directions agree to within 0.35 percentage points on every row
that moves, the floor is at most 0.30% at p50, and each measured delta lands
within 0.35 points of what removing one 1 ms request predicts.
`xlsx_range_source_first_cell` opens outside its timed region and does not move.
**`xls_range_source_open` is an OLE2 selector that never enters the ZIP locator**;
its +1.52% in one direction against −0.09% in the other comes from a transient in
the B1 leg (mean +2.16%, p95 +4.52%) and is reported rather than smoothed. A
sixth case, `xlsx_range_source_list_sheets`, issues zero requests and runs in
about 0.1 µs; it is retained in the leg files and read as noise, not as a result.

*(b) Real packages on the same transport*, 60 samples and 10 warmups per fixture
per leg (`results/change-0632/workbook-timing/`), because the harness's range
corpora are synthetic packages with five members:

| package | requests | before p50 | after p50 | A1→B1 p50 | A2→B2 p50 | saving | floor A3→A4 | floor B1→B2 |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| `ConditionalFormattingSamples.xlsx` | 10 → **9** | 11,338.9 µs | 10,265.4 µs | **−9.47%** | **−9.46%** | **−1.07 ms** | +0.20% | −0.15% |
| `shapes.pptx` | 6 → **5** | 6,630.9 | 5,566.6 | **−16.05%** | **−16.07%** | **−1.06 ms** | +0.01% | +0.02% |
| `shape-soft-edges.pptx` | 8 → **7** | 8,824.2 | 7,764.5 | **−12.01%** | **−12.09%** | **−1.06 ms** | −0.31% | +0.06% |
| `comment.docx` | 6 → **5** | 6,389.2 | 5,332.7 | **−16.54%** | **−16.54%** | **−1.06 ms** | +0.09% | −0.02% |

The two paired directions agree to within 0.08 percentage points on every row,
every floor is at most 0.31%, and every saving is 1.06–1.07 ms — one request of
this transport, to the microsecond, on every fixture.

**Stated against change 0623, as the brief asks.** 0623 measured **−76.58%** on
the workbook open, taking it from 45 requests to 10. This record takes that same
open from **10 to 9** and measures a further **−9.47%**, 1.07 ms. Against change
0611's 45-request baseline the two changes together are 45 → 9: a measured
48,208.6 µs → 10,265.4 µs, **−78.7%**, on the same fixture and the same
transport. The remaining nine requests are one EOCD probe, one central-directory
read, four coalesced structural runs and three single-member reads.

**Local, in-process sources**, where a request costs about a microsecond rather
than a millisecond and **no claim is made either way**
(`results/change-0632/timing/local-summary.txt`):

| case | A1→B1 p50 | A2→B2 p50 | floor A3→A4 | floor A1→A2 | floor B1→B2 |
| --- | ---: | ---: | ---: | ---: | ---: |
| `opc_file_source_open` | −6.55% | −3.12% | −1.02% | +1.67% | **+5.41%** |
| `docx_file_source_open` | −0.56% | +4.91% | +0.92% | **−6.23%** | −1.07% |
| `pptx_file_source_open` | +2.68% | +3.32% | **−4.04%** | −3.02% | −2.42% |
| `docx_file_source_full_text` | **+8.55%** | +4.21% | −2.09% | −3.16% | **−7.03%** |

`docx_file_source_full_text`'s +8.55% exceeds the programme's 5% review trigger,
so it is reported here rather than folded into a mean, and it was chased rather
than explained away. Two legs of the **same after binary** differ by −7.03% at
p50 on that case in this window, and the two paired directions differ from each
other by 4.3 points, so the case cannot resolve an effect of that size on this
host at this load; the same is true of `opc_file_source_open` in the other
direction, where the B/B floor is +5.41%. What the counts predict locally is one fewer
`pread` of 46 bytes per open, and one 64 KiB allocation and zero-fill replaced by
one of `central_directory_size` — between 55 and 65 KiB of `memset` that no
longer happens. That is the direction `opc_file_source_open` — the only case in
the table whose timed region is an open and nothing else — measures, at −6.55%
and −3.12%. Nothing local is claimed.

## Correctness evidence

**The differential is the gate, and it is clean.** Change 0587 names change
0582's differential as the gate for any read-grammar change in `soapberry-zip`;
change 0611 extended it with the ordinary indexed read path. Run in full — both
builds, 22,875 inputs, both limit profiles, both directions, 2,886,786 member
verdicts — the two reports are byte-identical.

**The open oracle over 533 containers** compares the full source-backed OPC open
field by field, and differs only in the request count.

**Tests added.** `crates/soapberry-zip/tests/directory_prefill.rs`, 11 tests:

* `one_request_serves_the_probe_and_the_whole_directory_scan` — the index build
  costs exactly two requests, the second one exactly the declared directory at
  the declared offset.
* `the_prefill_finds_exactly_what_a_separate_scan_finds` — for 0, 1, 2, 17 and 64
  members, the index's members equal what the untouched public
  locate-then-`entries()` path finds.
* `a_directory_larger_than_the_window_is_read_in_two_requests` — a 1,600-member
  archive with a 100,800-byte directory: the whole open costs exactly three
  requests (the 22-byte EOCD probe, the 64 KiB window, one read for the rest),
  the second read continues forward from exactly where the window ends with no
  overlap and no gap, and the reads end exactly at the directory end.
* `a_zip64_archive_keeps_its_records_and_its_members` — both retained ZIP64
  fixtures with a real ZIP64 end-of-central-directory record.
* `a_comment_bearing_archive_keeps_the_backwards_search_and_its_members` — a
  comment pushes the EOCD off `len - 22`; the directory is still read exactly
  once and the separate 46-byte probe is gone.
* `a_comment_that_contains_an_eocd_signature_resolves_where_it_always_did` — the
  false-signature case the locator documents.
* `a_prefixed_archive_still_reaches_the_base_offset_fallback` — 4,096 bytes of
  prefix: both probes happen, at the declared offset and at
  `eocd - central_dir_size`, and the members match the unprefixed archive.
* `a_short_reading_source_is_served_the_same_members` — a source capped at 1, 7,
  46, 64, 97 and 512 bytes per call.
* `a_directory_shorter_than_one_central_record_takes_no_prefill` — no read longer
  than 46 bytes is issued at the directory offset, and the typed `Eof` stands.
* `an_oversized_variable_field_keeps_its_typed_refusal` — the spill-boundary pin,
  compared directly against the untouched 64 KiB public path.
* `an_empty_archive_takes_no_prefill_and_opens`.

**Tests updated.** Three request-count assertions in change 0623's
`crates/litchi-opc/tests/structural_prefetch.rs`, from 4, 10 and 14 reads to 3,
9 and 13. They are the only assertions in the repository that pin an absolute
ZIP open request count, and each falls by exactly the one locator request this
change removes. No test was weakened: each still pins an exact number, and
`a_managed_open_keeps_the_exact_grammar` still asserts that the managed open
costs strictly more reads than the unmanaged one.

**Gates.** Twelve sections, all exit 0: formatting; Clippy for
`soapberry-zip --all-targets` with workspace lints denied (the crate has no
features, so `--all-features` is the same run); rustdoc; the `soapberry-zip`
suite; and the suite of **every crate that depends on `soapberry-zip`** —
`litchi-opc` (also `--all-features`), `litchi-xlsx`, `litchi-docx`,
`litchi-pptx`, `litchi-odt`, `litchi-odf-common`, `litchi-odc`, `litchi-odg`,
`litchi-odp`, `litchi-oth`, `litchi-odf-formula`, `litchi-ppt`, `litchi-sign`
and `litchi-iwa-archive` — plus `litchi-ooxml-common`, `litchi-core`, the
`litchi` facade with `--features docx,xlsx,pptx,xls`, and the harness's own
suite in `tools/perf-baseline`. The last two are change 0587's two gaps,
reachable from no per-crate gate. Every gate tail is in
`results/change-0632/gates.txt`.

## Validation preserved

No validation moved, weakened or changed identity. The EOCD fast-path gate, the
backwards search, `max_search_space`, ZIP64 locator and record resolution,
`validate_classic_single_disk`, `EndOfCentralDirectory::create`,
`physical_entry_bound`, every `ArchiveLimits` ceiling, the member-name and
metadata-byte charges and the oversized-record spill all run in the same order,
on the same bytes, with the same results — which the byte-identical differential
report measures rather than asserts. No `unsafe` was added, no defence against
malformed input was relaxed, no limit was raised, and no typed refusal was traded
for a partial result. The one difference a caller can observe is that its
`read_at` is called once fewer per index build, with 46 fewer bytes.

## Limitations

**No claim is registered.** `performance_claim: none`. The figures above are
evidence.

**The latency-bearing result is a simulated transport**, a fixed service time per
physical request, not a network or a real device. What it prices is the request
count, and the request count is measured deterministically. No cold page cache,
real device, peak-RSS, instruction-count, syscall, concurrency-scaling or
cross-platform result is claimed.

**Nothing is claimed on a local, in-process source.** The local table is reported
because it was run; its A/A and B/B floors are larger than every delta in it.

**The window is measured against this repository's corpus.** 64 KiB against a
largest observed central directory of 9,724 bytes. A producer writing a directory
larger than 64 KiB gets the head of it in the window and the rest in the reads
the scan issues today — never an incorrect result, and never more requests than
today.

**The locator's own 64 KiB scratch stays.** Its size is what holds the backwards
EOCD search to one request for a comment-bearing or trailing-byte archive, and
the directory size is not known before the locate. A two-stage locate — a small
stack buffer for the `len - 22` fast path, falling back to the heap buffer for
the 1 container in 533 that needs the search — would remove that allocation too;
it is not attempted here because it costs one extra request on every archive that
does not take the fast path and the saving was not measured.

**`PreservationIndex` still scans the directory for itself.** It takes a caller
scratch and is reached after the locate has returned, so it gets no prefill. Its
request count is unchanged.

**A managed open gains the same request, and consumes 46 bytes less budget.**
This change is below the `litchi-opc` layer entirely, so it applies on the
managed path too — change 0623's `a_managed_open_keeps_the_exact_grammar`
measures 14 reads falling to 13. That is a movement in *when* a finite
`Resource::InputBytes` budget is exhausted, and it is stated rather than
claimed away: a managed open now reserves and charges `central_directory_size`
once where it reserved and charged 46 and then `central_directory_size`, and it
makes one fewer `ExecutionContext::check` cancellation observation. The movement
is exactly change 0611's — the merged read *is* the read the scan had to issue,
one for one, with a strict subset removed — and it is in the direction of less
work: an open whose budget sat within 46 bytes of the boundary now succeeds
where it refused, and none can refuse where it succeeded. It is **not** change
0623's case, where a speculative read belonging to no member was reserved as a
unit; that is why 0623 had to gate its mechanism off the managed path and this
one does not. No managed timing measurement is offered.

## Retained evidence

Evidence packet, including the census, the count probe, the open oracle, the full
read-grammar differential, every timing leg and the gate tails:
[`results/change-0632/README.md`](results/change-0632/README.md).
