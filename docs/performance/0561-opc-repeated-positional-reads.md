# 0561: repeated positional reads in file-backed OPC

Status: attribution only. No production change and `performance_claim: none`.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was measured

Five perf-harness children, each `--warmup 0 --samples 1`, traced with
`strace -f -e trace=pread64` while reading a file-backed OOXML package:

| Capture | `pread64` calls | distinct (offset, length) ranges | calls that re-read a range | immediate back-to-back duplicates | mean read |
| --- | ---: | ---: | ---: | ---: | ---: |
| `docx_file_source_open` | 266 | 23 | 91.4% | 22.9% | 219.7 B |
| `docx_file_source_full_text` | 282 | 23 | 91.8% | 23.0% | 228.3 B |
| `pptx_file_source_open` | 13,592 | 1,262 | 90.7% | 25.0% | 120.6 B |
| `pptx_file_source_list_slides` | 16,792 | 1,262 | 92.5% | 25.0% | 125.6 B |
| `pptx_file_source_selected_slide` | 13,608 | 1,262 | 90.7% | 25.0% | 120.6 B |

About nine in ten positional reads re-read bytes the same child already read, and
a quarter repeat the immediately preceding range verbatim. Individual ranges are
read 6, 10, 12, 20 and in one case 24 times. Reads are small: for
`pptx_file_source_selected_slide`, 6,760 calls are 16 bytes or fewer and 3,400
are 32 or fewer.

These filesystem selectors use a fresh child per sample, so differencing
`strace -f -c` runs at 5 and at 25 samples and dividing by 20 isolates one
complete child: its untimed open, the timed operation, and the post-timer
oracle. It does **not** isolate the timed operation alone. On that basis
`docx_file_source_full_text` costs 598 `statx` and 262 `pread64` per child and
`pptx_file_source_selected_slide` costs 9,696 and 13,588. The owned-source
selectors `xlsx_source_first_cell` and `opc_source_open_main_read` isolate to
zero of both, which confirms two things: the counts come from the file adapter
rather than from the harness, and process start-up contributes no `statx` or
`pread64` of its own.

`pptx_file_source_open` accounts for almost all of the calls that
`pptx_file_source_selected_slide` makes — 13,592 against 13,608 — so the
repetition is in the open and index path, not in the slide query that the
selected-slide selector actually times.

## Where the reads come from

Every member read on this path costs exactly four `pread64` calls, because
`litchi-opc` reaches `soapberry_zip::office::IndexedArchive::read_entry`, which
builds a fresh read session per call:

| # | Read | Issued by | Size |
| ---: | --- | --- | ---: |
| 1 | local file header at `local_header_offset` | `ZipArchive::get_entry`, `archive.rs:1636` | 30 B |
| 2 | compressed payload at `body_offset` | `RangeReader::read` under flate2's 32 KiB `BufReader` | member-sized |
| 3 | data descriptor at `data_end` | `ZipVerifier::read`, `archive.rs:1982` | 16 B |
| 4 | **the same data descriptor again** | the same branch, fired by the terminating `Ok(0)` of `read_to_end` | 16 B |

The model is arithmetically exact against the histogram: 3,400 reads of 32 bytes
or fewer are the header reads, so 3,400 member reads; 6,760 reads of 16 bytes or
fewer is 1.988 descriptor reads per member read; the remaining 3,448 are payload
reads; and 3,399 immediate duplicates is one per member read.

Two independent causes drive the 6x-to-20x multiplicity, and neither is a loop
inside one library operation:

- structural members bypass the part cache. `read_structural_member`
  (`litchi-opc/src/pkgreader.rs:74-87`) calls `IndexedArchive::read` directly, so
  `[Content_Types].xml` and every `.rels` is re-read in full on every package
  open;
- the harness child itself opens the package twice — once untimed to prepare and
  once in the post-timer oracle — and the oracle then sweeps all 200 slides.

So the multiplicity is process-tree and oracle work, not one operation looping.
The per-read waste is the library-side defect.

## What is already in memory

`IndexedArchive` retains, from its single central-directory scan, each entry's
name, flags, compression method, sizes, **CRC**, `local_header_offset` and
descriptor presence. Two of the four reads recover values that cannot differ
from what is retained:

- the descriptor read's CRC is provably equal to the retained `crc`, because
  `DataDescriptor::parse_complete_with_width` only returns successfully when the
  descriptor matches the entry's CRC and both sizes. The read is a physical
  consistency check, performed twice per read on every read: 6,760 of 13,608
  calls, 49.7%;
- the local-header framing is invariant. `get_entry` re-derives `body_offset`
  from a fresh 30-byte read each time, although the `ReaderAt` contract requires
  the source to be byte-stable for the reader's lifetime: 3,400 calls, 25.0%.

About 73% of the positional reads on this path therefore recover values already
held in memory or re-derive an invariant.

## Ownership

ADR 0011 makes `litchi-opc` the only owner that translates between OPC
operations and the selected ZIP implementation, and ADR 0010 keeps archive types
out of the facade. Changes to reads-per-entry-read belong in `soapberry-zip`,
which owns the ZIP grammar; changes to how many entry reads happen, or to the
read policy, belong in `litchi-opc`. `litchi-pptx` and the facade own nothing
here.

## Follow-up

Change [0562](0562-zip-descriptor-read-once.md) removes the duplicate descriptor
read. Memoizing the resolved descriptor and the local-header framing per entry,
coalescing the header/payload/descriptor span for small members, and giving
structural members a retention policy remain open and unmeasured.

After 0562, the `docx_file_source_open` capture's 224 remaining calls cover 23
distinct ranges. Counting repeats names the next targets precisely: one
840-byte span at offset 64 read 24 times, a 4-byte probe at offset 0 read 20
times alongside the first local header at the same offset read 20 times, an
8-byte probe at offset 30 read 10 times, and the end-of-central-directory and
central-directory spans read 10 times each. The child performs about ten
independent package opens, so most of that is per-open re-work that no
within-open cache can remove; the within-open residual is the roughly 2.4 reads
of the structural member and 2 reads of the first local header per open.

The PPTX capture shows the same structure at scale. After 0562 its 10,216 calls
group as 3,386 reads of 30 bytes, 3,376 of 16 bytes and 3,414 payload reads —
one header, one descriptor and one payload per member read. Memoizing the
resolved descriptor and the header framing per entry would therefore remove
about two thirds of what remains. That is a count model derived from the
retained trace, not a measured result, and neither change has been implemented
or measured.

## Limitations

No production change, latency, resource or cold-cache measurement is claimed
here — only that the reads occur. The captures are warm-cache and single-sample;
they count syscalls, not physical device I/O.

[Evidence, traces and the replay command](results/change-0561/README.md).
