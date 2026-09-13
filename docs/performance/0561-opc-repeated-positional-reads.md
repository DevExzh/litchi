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
read that occurs *within* one verified entry read.

### Correction: where the remaining repeats actually are

This record originally projected that memoizing the resolved descriptor and the
local-header framing per entry would remove about two thirds of the remaining
reads. **That projection was wrong**, and it is retracted here.

It assumed the 6x-to-20x range multiplicity happens inside one archive instance.
It does not. Segmenting each traced `(pid, fd)` read stream at its own 22-byte
end-of-central-directory read — the unambiguous marker for one archive
construction — separates the two cases:

| Capture | `pread64` | archive constructions | local-header reads | repeats *within* one archive | descriptor reads | repeats within one archive |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `docx_file_source_open` | 224 | 10 | 52 | 9 | 42 | **0** |
| `docx_file_source_full_text` | 236 | 10 | 56 | 9 | 46 | **0** |
| `pptx_file_source_open` | 10,216 | 10 | 3,386 | 9 | 3,376 | **0** |
| `pptx_file_source_selected_slide` | 10,228 | 10 | 3,390 | 9 | 3,380 | **0** |

Inside one `IndexedArchive`, every member's local header and descriptor is
already read exactly once. The nine header repeats are all `(offset 0, length
30)` — a prologue probe taken before any archive exists, not a `get_entry`
repeat. So a per-entry memo of either value would remove **zero** reads on this
corpus. It would still be correct, and would pay off whenever the OPC part cache
evicts or bypasses and a member is read twice from one archive, but it is not
what this measurement calls for.

Reproduce with
[`segment_reads.py`](results/change-0561/segment_reads.py); the result is in
[`read-segmentation.json`](results/change-0561/read-segmentation.json).

### What the measurement does call for

Every member read costs three positional reads after 0562 — a 30-byte local
header, the payload, and a 16-byte descriptor — and each is a separate `pread64`
on the **first** read of that member, which is what this corpus does. Two
targets follow:

1. **Coalesce the three into one bounded read per member.** The span is
   `local_header_offset` to `body_end + descriptor`, bounded by the central
   record's `compressed_size` plus a fixed metadata allowance. The payload size
   distribution supports it: 2,136 of the roughly 3,414 PPTX payload reads are
   256 bytes or fewer and about 3,394 are 2 KiB or fewer, so a small window
   covers nearly every member. This is the only change measured here that would
   reduce first-read cost.
2. **Stop reconstructing the package per open** — ten archive constructions per
   traced child, each re-reading the end-of-central-directory record, the
   central directory, the detection probes and the structural members. ADR 0011
   makes this `litchi-opc`'s to own, not `soapberry-zip`'s.

Neither is implemented or measured. The counts above are properties of the
retained traces; the coalescing figure is a span model, not a result.

## Limitations

No production change, latency, resource or cold-cache measurement is claimed
here — only that the reads occur. The captures are warm-cache and single-sample;
they count syscalls, not physical device I/O.

[Evidence, traces and the replay command](results/change-0561/README.md).
