# 0564: 95% of a source-backed XLS open is a four-byte header scan

Status: attribution only. No production change and `performance_claim: none`.
It also records a conflict with four tested contracts, and corrects an inference
in change [0325](changes/0325-cfb-frame-transaction-rejected.md) that closed this
line of work.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was measured

One source-backed XLS open over `test-data/ole/xls/ConditionalFormattingSamples.xls`
(1,402,368 bytes) on a `litchi_core::FileSource`. Counts are isolated by
differencing `strace -f -c` runs at 1 and 11 samples with one warmup and dividing
by ten, so process start-up and staging cancel:

| | per open |
| --- | ---: |
| `pread64` | **655** |
| `statx` | **636** |

A full `strace -f -e trace=pread64` capture of the same child gives the shape:

| request size | reads (2 opens) | share |
| ---: | ---: | ---: |
| **4 bytes** | **1,242** | **94.7%** |
| 512 bytes | 48 | 3.7% |
| 65,536 bytes | 14 | 1.1% |
| 49,152 / 43,473 / 1,536 / 840 | 8 | 0.6% |

Only 6 of 1,311 consecutive reads continue where the previous one ended, because
each four-byte read is followed by a payload skip that performs no read.

## Where the 655 reads are

An independent static model reconstructs every call from the code and the
fixture's CFB and BIFF geometry. It reproduces the measured counts **and** the
logical byte totals recorded in change 0325 exactly — 655 reads / 567,685 bytes
for open, 921 / 569,398 for one-cell — so the attribution below is confirmed
rather than estimated.

| Group | Reads | Bytes | Shape | Site |
| --- | ---: | ---: | --- | --- |
| **A** BIFF globals header pre-pass | **621** | 2,484 | one 4-byte read per BIFF record | `litchi-xls/src/workbook/source.rs:1515` |
| B globals bulk range read | 9 | 551,377 | one read per contiguous FAT run | `source.rs:1582` → `litchi-cfb/src/shared.rs:904` |
| C CFB FAT sectors | 22 | 11,264 | one 512-byte read per FAT sector | `litchi-cfb/src/file.rs:884` |
| D CFB header | 1 | 512 | one-shot | `file.rs:590` |
| E CFB directory | 1 | 1,536 | one-shot, already batched | `file.rs:995` |
| F CFB MiniFAT | 1 | 512 | one read per MiniFAT sector | `file.rs:955` |

The CFB structural parse is 25 reads. The BIFF globals scan is 630. Everything
else — `select_workbook_stream`, `stream_len`, `stream_cursor_at`,
`worksheet_count`, `worksheet_names` and every `ensure_current_parts` — resolves
the in-memory index and reads nothing, which is why open and list have identical
read counts.

**Group A is 94.8% of the reads and 0.44% of the bytes.** Group B then reads
`[0, global_end)` in full, a strict superset of every byte group A read, so
those 2,484 bytes are read twice.

For the one-cell query the extra 266 reads are 177 worksheet record headers, 88
payload reads and one 7-byte SST entry fetch. A skipped record costs exactly one
read; a consumed record costs exactly two.

## Correcting change 0325

Change 0325 rejected a bounded frame prefetch for this path and closed the line,
stating that `xls_source_backed_open` and `xls_source_backed_open_list_worksheets`
"have no `WorksheetScan` opportunity" and are "therefore zero-opportunity cases".

The first half is literally correct: **zero** of open's 655 reads come from
`WorksheetScan`. The inference does not follow. 621 of them — 95% — come from a
different loop, `parse_globals`'s header pre-pass, which change 0325 never
examines. The candidate it evaluated could have addressed at most 265 of the 921
one-cell reads and none of open's 655, and the line was closed on that basis
while the dominant group went unmeasured.

Change 0325's version-call figures (1,266 and 1,813) also no longer describe
HEAD; changes 0558 and 0560 halved them to 631 and 902. Its read and byte
figures are unchanged and still exact.

## The conflict

Fusing the two passes would take open from 655 reads to about 34 and from 631
source observations to about 13, and would stop reading 2,484 bytes twice. At
roughly 116 ns per syscall — derived from the measured 149 microsecond
file-source penalty over 1,286 syscalls — that is about 144 microseconds, which
is essentially the whole penalty.

It cannot be done without changing four contracts that tests assert deliberately:

| Contract | Test |
| --- | --- |
| Open's range list is *exactly* one 4-byte range per global header, then one bulk range | `raw_handoff_reads_global_headers_then_one_exact_global_range`, `litchi-xls/tests/source_backed.rs:1509` |
| A `FILEPASS` payload is never read | `filepass_header_scan_never_reads_its_payload_or_worksheets`, `:1392` |
| A skipped payload is never read | `supported_unknown_payload_is_skipped_without_source_overread`, `:1252` |
| Open never touches a sheet body | `selected_queries_are_bounded_to_the_selected_owner`, `:949` |

The source comment at `source.rs:1502-1504` states the intent directly: the pass
reads only headers so that `FILEPASS` is detected before any payload byte is
read, and so the boundary is exact before the bulk read.

A fused pass would therefore have to establish, explicitly:

1. a "buffered but never published" rule for payload bytes held in a chunk,
   including a `FILEPASS` payload;
2. a bound on reading past `global_end`, which is unknown until framing
   completes, so the last chunk spills into the first sheet body. The
   `BoundSheet8` records inside globals carry each sheet's stream position and
   every sheet start is at or above `global_end`, so clamping fills to the
   smallest start seen is sound — but it only becomes available part-way through
   the scan;
3. how `max_global_bytes` and `max_global_records`, checked per record before
   more bytes are read today, fold into a chunk size;
4. that collapsing 621 freshness observations to about 9 is an accepted
   reduction in change-detection granularity.

That is precisely the "explicit contract for distinguishing bytes that are
payload from bytes that may be skipped" change 0325 said was missing. This
record does not establish it. Doing so changes deliberate, tested behaviour and
is a decision to take explicitly rather than as a side effect of an
optimization.

## Also recorded

`load_fat` and `load_minifat` (`litchi-cfb/src/file.rs:884`, `:955`) read one
sector per call through `read_sector_into`, while `read_sectors_batched`
(`file.rs:2016`) sits in the same file and is already used by `load_directory`.
This fixture's 22 FAT sectors are spread about 129 sectors apart, so batching
recovers nothing here; it would help a file with a contiguous FAT. The CFB
header and FAT sector 0 are physically adjacent and read separately, worth one
read.

## Limitations

No production change, latency, resource or cold-cache measurement is claimed.
The syscall counts are warm-cache and from this host; the attribution model is
static, validated against the measured totals rather than against an
instrumented run. One fixture and one query are covered; straddle-split reads
are zero here and would be positive elsewhere.

[Traces, the isolation pair and the summarizer](results/change-0564/README.md).
