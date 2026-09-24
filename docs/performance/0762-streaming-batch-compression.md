# 0762 — batched owned Deflate: fixed 16 KiB chunks, no pre-finish sync flush and level 5 cut streaming DOCX creation by 28%, XLSX by 45% and PPTX by 10%, with deterministic, near-identical-size packages

Status: retained, implemented in `soapberry-zip`. `performance_claim: none` —
the paired timings, instruction counts, sizes and allocation counts below are
evidence, not a registered claim.

OLE2 and OOXML remain the active priority; ODF stays deferred and iWork is
excluded. Base `1d1044e3ac`; branch `perf/0762-streaming-compression-and-leases`,
commit `bc18e8abdd`. The coordinator's task: apply owner decision 4 of
[0758](0758-owner-decisions-2026-09-24.md) ("accept the batch compression
policy") to the streaming creation writers — coalesce their small writes before
the compressor at fixed, caller-independent boundaries, drop the per-member
sync flush, and evaluate per-member compression levels — without moving any
byte of the preservation paths.

**Result.** With both legs built by the identical command,
`docx_streaming_create` falls from 49.353 to 35.661 ms on the large corpus
(−27.77%, 95% CI [−28.47%, −27.22%]) and from 3.156 to 2.278 ms on the medium
one (−27.84%); `xlsx_streaming_create` from 163.942 to 90.550 ms (−44.81%) and
from 10.631 to 5.639 ms (−47.01%); `pptx_streaming_create` from 190.417 to
171.484 ms (−9.89%) and from 6.378 to 5.835 ms (−8.57%). The streaming
packages change bytes, as decision 4 allows: every member decompresses to
exactly the bytes it held before (17,050 members over nine corpora compared),
eight of the nine harness corpora get smaller (−0.07% to −2.17%) and the
64-row XLSX one grows by 4 bytes (+0.12%). The output is a pure function of
the member bytes: re-splitting the caller's writes, short and interrupting
sinks, and separate processes give identical archives. The two preservation
controls publish the same bytes and move by +1.19% (CI [−0.08%, +1.53%]) and
−3.50% (host-load noise, CI [−9.00%, +10.07%]).

## What was changed

All production code is in `crates/soapberry-zip/src/writer.rs`:

* **The staged protocol for owned Deflate entries.**
  `ReusableDeflateState` gains an input stage (`stage`, `stage_consumed`) and
  three owned-path methods:
  * `write_owned` accepts caller bytes into the member's current chunk, at most
    up to the next chunk boundary, after handing a chunk completed by an
    earlier call to the codec. An error always means the call accepted
    nothing; a refused chunk stays staged and a retry resumes it.
  * `compress_stage` hands the staged bytes the codec has not consumed to the
    codec, draining the 32 KiB output buffer completely before every codec
    call and after the last one, and empties the stage once it holds a whole
    chunk.
  * `flush_owned` (an explicit `Write::flush`) hands over the staged bytes,
    then runs the existing sync flush and sink flush; `finish_owned` hands over
    the staged bytes, flushes the sink, and runs `Finish` — with no sync flush
    before it.

  The chunk is the frozen format constant `OWNED_DEFLATE_STAGE_BYTES`
  (16 KiB), cut at absolute member offsets `k × 16384`. The stage is allocated
  once per archive, on the first owned Deflate write, with `try_reserve_exact`;
  an allocation failure fails the write (`io::ErrorKind::OutOfMemory`) instead
  of falling back to an unstaged stream, because the chunk boundaries are part
  of the output.
* **Levels.** `OWNED_DEFLATE_LEVEL = 5` for owned members;
  `DEFAULT_DEFLATE_LEVEL = 6` for every borrowed member. `begin_member` (the
  borrowed paths) and the new `begin_owned_member` both go through
  `begin_member_at(level)`, which after `Compress::reset` calls
  `Compress::set_level` only when the level differs (zlib-rs changes only the
  matcher parameters on a reset stream; a refusal would fall back to a fresh
  codec at that level). `ZipArchiveWriter::take_owned_deflate` is the owned
  counterpart of `take_reusable_deflate`.
* **The owned entry's finish.** `ZipOwnedEntryWriter::finish` takes the
  descriptor through the new unflushed `ZipDataWriter::into_parts_with_crc_stage`
  and lets `OwnedCompressor::finish` flush: a Store entry flushes its sink
  exactly as its `flush` does; a Deflate entry ends through `finish_owned`.
  `ZipDataWriter::finish` (public, used by borrowed entries) still flushes and
  then takes the parts.
* `OwnedCompressor::write` and `flush` call `write_owned` and `flush_owned`
  for Deflate, with the same archive poisoning on a non-interrupted error.
* The shared codec call is factored into `run_codec`, used by `compress_once`
  and `compress_stage`.
* `ReusableDeflateState`'s `Debug` prints counts (level, pending, staged,
  consumed) instead of deriving over its buffers, so no payload or compressed
  byte reaches a debug string.

Documentation on `ZipArchiveWriter::start_file_owned`, `ZipOwnedEntryWriter`
(including `compressed_bytes` and `finish`) and
`office::StreamingArchiveEntry` states the batching, the determinism and the
error timing below.

**Who is affected.** In production, `start_file_owned*` is reached only
through `StreamingArchiveWriter::start_entry` → `PhysPkgWriter::start_part` /
`start_stored_part`, whose only production callers are the streaming DOCX,
XLSX and PPTX writers (`crates/litchi-{docx,xlsx}/src/streaming.rs`,
`crates/litchi-pptx/src/writer/streaming.rs`). The preservation writer
(`preserve.rs`, `ReusedDeflateEncoder`), `write_deflated`,
`write_deflated_stream`, `write_deflated_sized`, `PhysPkgWriter::write` and
therefore every regenerated or source-backed member of an opened package,
every fresh non-streaming OOXML writer and every ODF writer keep level 6, one
codec call per write, the pre-finish sync flush and today's bytes. No ODF or
iWork crate reaches the owned path.

**Breaking changes** (alpha; standing trade-off 1 of
[0652](0652-owner-decisions-for-the-third-wave.md)):

* the bytes of every owned Deflate member, and so of every package the three
  streaming writers create;
* an owned Deflate entry's `Write::write` accepts at most up to the next 16 KiB
  boundary (a short write; `write_all` callers see no difference);
* compressed output, a sink failure and a compressed-size or output-limit
  refusal can surface up to one chunk of input later than before (below);
* one more allocation of 16 KiB per archive that has an owned Deflate member;
* `Debug` of `ZipOwnedEntryWriter` and `StreamingArchiveEntry` shows the
  state's counts instead of its buffers.

Tests: in `crates/soapberry-zip/tests/deflate_reuse.rs` the five tests that
pinned owned output to a fresh per-write `DeflateEncoder` are replaced (listed
under *Verification*), and `crates/litchi-docx/tests/streaming.rs` gains one.

## Authority

Owner decision 4 of [0758](0758-owner-decisions-2026-09-24.md): creation-from-
scratch writers may produce new output bytes, by coalescing small writes
before the compressor or by cheaper per-member compression strategies;
preservation of existing packages is unaffected. Standing trade-offs of
[0652](0652-owner-decisions-for-the-third-wave.md): 1 (breaking changes are
acceptable; the ones above are documented), 2 (correctness first: every limit
is still enforced before output, a stage that cannot be allocated fails
closed), 3 (the benign common path — many small writes — is the one made
cheaper); decision 10 of 0652 prefers smaller files, which the level choice
below weighs. ADR 0006 (deterministic serialization: the output is a function
of the member bytes, proven below; preservation is the default and its bytes do
not move), ADR 0005 (bounded streaming: the stage is a fixed 16 KiB transport
buffer outside the semantic accounting, like the 32 KiB output buffer and the
4 KiB CRC stage of [0752](0752-streaming-writer-small-write-batching.md)),
ADR 0010, 0011 and 0024 (the change lives in the archive owner; no archive type
crosses into a format crate; no dependency edge is added).

## The evidence that motivated it

* [0752](0752-streaming-writer-small-write-batching.md) measured that zlib-rs
  output depends on how its input is split into calls, which is why 0752 staged
  only the CRC-32. Its probe prices the Deflate call overhead on the DOCX
  `word/document.xml` payload at 21.76 ms per caller write against 11.33 ms in
  16 KiB chunks, with a slightly smaller stream (376,031 → 375,255 bytes;
  `results/change-0752/probe/crc_deflate_probe-output.txt`).
* The coordinator's timed-region profile of the base
  (`results/change-0756/profile-r2/REPORT.md`) attributes 73–76% of the XLSX
  large iteration to Deflate (`longest_match` 53%), and 63% of the PPTX large
  one to Deflate over 16,421 small members: dynamic Huffman construction at
  block flush 42% inclusive, the per-member reset 6.4%, `memset` 10%.

Why the old output depended on the caller's writes: the owned entry made one
codec call per caller write. zlib-rs's `deflate_medium` keeps its pending
match in locals that a call returning for more input discards, so where the
input stops is part of the stream. A fixed chunking at absolute offsets makes
the call sequence a function of the member bytes.

## Why the output is a pure function of the member bytes

A codec call's result depends only on the codec state, the input slice, the
output space and the flush mode. Under the staged protocol:

* a codec call is made only by `compress_stage` (with flush `None`) or by the
  existing flush and finish loops, and only after the output buffer has been
  drained completely, so every call has the whole 32 KiB buffer;
* the input of a `None` call is always "the rest of the current chunk",
  delivered only once the chunk is complete (absolute offset `k × 16384`), or
  at an explicit flush or the finish;
* the drains, short sink writes and retried interruptions happen between codec
  calls and never change their arguments.

By induction on the calls, the stream is determined by the member bytes and
the offsets of explicit flushes. The staged model in the tests
(`StagedModel`, written against flate2's raw `Compress`) states this protocol;
flate2's own `DeflateEncoder` fed the same 16 KiB pieces and finished without a
flush agrees with the model byte for byte on patterned, random and empty
payloads, and on the chunk boundaries themselves.

## Choosing 16 KiB, dropping the sync flush, and the level

`level_probe` (source in `results/change-0762/probe/level_probe/`) builds the
harness's streaming corpora with the real writers, extracts every member, and
recompresses the members under the staged protocol with one reused codec per
corpus (reset, then `set_level`), timing whole passes on CPU 12 (minimum of 8
passes for large corpora, 30 for medium). Sizes are exact; "archive" adds the
corpus's unchanged ZIP framing.

**Chunk size and sync flush (level 6):**

| corpus | 4 KiB | 16 KiB | 64 KiB | one call | 16 KiB + sync flush |
| --- | ---: | ---: | ---: | ---: | ---: |
| DOCX large member data, bytes | 375,682 | 375,662 | 375,656 | 375,657 | 375,680 |
| XLSX large member data, bytes | 2,572,953 | 2,558,832 | 2,555,879 | 2,556,149 | 2,558,869 |
| PPTX large member data, bytes | 5,346,126 | 5,346,108 | 5,346,100 | 5,346,100 | 5,457,493 |
| XLSX large compression, ms | 125.8 | 125.3 | 125.7 | 123.6 | 124.7 |
| PPTX large compression, ms | 129.2 | 130.6 | 130.5 | 128.6 | 133.6 |

Once the codec gets at least 4 KiB per call, call overhead no longer matters;
16 KiB is within 0.11% of one call in size and bounds the stage at a quarter of
the output buffer. The pre-finish sync flush costs 6.8 bytes per member (111 KB
on the 16,421-member PPTX corpus) and 3 ms of PPTX compression.

**Levels 1–6 under the staged protocol, harness corpora** (archive bytes and
compression ms; base = today's archive):

| level | DOCX large | XLSX large | XLSX medium | PPTX large | PPTX medium |
| --- | --- | --- | --- | --- | --- |
| base | 376,848 | 2,563,433 | 167,418 | 7,940,406 | 274,398 |
| 1 (`deflate_quick`) | 604,563 · 4.6 | 4,057,185 · 16.6 | 255,710 · 0.94 | 9,211,470 · 39.7 | 320,558 · 1.33 |
| 2 (`deflate_fast`) | 378,456 · 6.2 | 2,595,646 · 24.4 | 165,826 · 1.43 | 7,898,801 · 116.6 | 274,352 · 4.20 |
| 3 | 378,416 · 6.3 | 2,557,118 · 43.8 | 163,684 · 2.44 | 7,823,290 · 122.7 | 271,284 · 4.42 |
| 4 | 376,230 · 13.6 | 2,556,882 · 58.8 | 163,719 · 3.40 | 7,822,041 · 125.0 | 270,864 · 4.49 |
| **5** | **376,054 · 13.9** | **2,559,706 · 64.8** | **163,790 · 3.80** | **7,821,547 · 129.2** | **270,370 · 4.55** |
| 6 | 376,054 · 14.1 | 2,559,622 · 125.3 | 163,800 · 6.21 | 7,821,514 · 130.6 | 270,349 · 4.67 |

**Levels on real content.** The same protocol over every member of every
workspace fixture that opens (62 DOCX, 180 XLSX and 78 PPTX packages from
`test-data/`; 959, 3,011 and 3,417 members), with "base6" the old protocol
for a member written in one call (one codec call, then the sync flush), at
level 6 (`probe/level-real.tsv`):

| format | base6 | level 6 | level 5 | level 4 | level 3 | level 1 |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DOCX bytes | 4,050,834 | −0.15% | **+0.06%** | +1.19% | +2.22% | +22.2% |
| XLSX bytes | 4,299,065 | −0.44% | **−0.24%** | +1.28% | +3.93% | +32.0% |
| PPTX bytes | 4,466,528 | −0.50% | **−0.27%** | +1.07% | +2.21% | +24.2% |
| DOCX ms | 95.3 | 94.9 | **88.2** | 81.3 | 74.6 | 33.2 |
| XLSX ms | 143.3 | 142.6 | **111.2** | 100.1 | 85.5 | 33.4 |
| PPTX ms | 108.7 | 108.3 | **99.8** | 91.7 | 86.0 | 34.4 |

Level 5 is the only level that saves material time without material growth.
It ends zlib-rs's `deflate_medium` match search at a chain of 32 and a match of
32 bytes instead of 128 and 128. Against level 6 under the same protocol its
growth is +0.21%, +0.21% and +0.23% on the real members and ≤ 0.01% on the
harness corpora (0 on DOCX, whose level-5 and level-6 archives are identical),
and it halves the XLSX sheet's compression (125 → 65 ms) where level 6's long
chains find nothing better. Against today's output it is smaller on every real
format but DOCX (+0.06%) and on eight of nine harness corpora. Levels 3 and 4
grow real content by more than 1% and are rejected; levels 1 and 2 grow it by
4.9–32%. The tiny members that dominate PPTX compress to the same size at levels
3–6: their cost is the per-member dynamic Huffman construction (levels 2–6) and
the hash-table reset, which no level short of 1's fixed codes avoids, and level
1 grows them by 25%. Level 6 remains the size-optimal alternative (the
real-content growth of 5 over 6 is the price of the XLSX saving); it is one
constant.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, shared with other agents
(load average 21–25 during the campaign); every measured process pinned to
CPU 12 with `taskset`; no build of this record ran during a measurement.
Harness `tools/perf-baseline`, unchanged, with its deterministic generators.
Both legs are built by the identical command from their own trees: before from
a detached worktree at `1d1044e3ac`, after from the branch at `bc18e8abdd`
(`binaries.txt`). The two binaries are copied to paths of equal length
(`bin/b/` and `bin/a/`), so argv and path lengths match.

### Timing

Four rounds; in each, every case ran before, after, after, before, so eight
processes per leg, paired within the round (slot 1 with 2, slot 4 with 3). The
table shows the medians of the process p50s and p95s, the median paired p50
change and a percentile bootstrap 95% interval over the eight paired changes
(20,000 resamples, seed 762). Raw reports: `timing/abba-raw.tar.gz`; summary:
`timing/summary.json`.

| case, corpus | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after |
| --- | --- | ---: | ---: | ---: | --- | --- |
| `docx_streaming_create` large | 8+8 × 20 | 49.353 | 35.661 | **−27.77%** | [−28.47%, −27.22%] | 50.123 → 35.914 |
| `docx_streaming_create` medium | 8+8 × 40 | 3.156 | 2.278 | **−27.84%** | [−29.23%, −27.27%] | 3.218 → 2.295 |
| `xlsx_streaming_create` large | 8+8 × 15 | 163.942 | 90.550 | **−44.81%** | [−45.20%, −44.55%] | 165.256 → 90.948 |
| `xlsx_streaming_create` medium | 8+8 × 30 | 10.631 | 5.639 | **−47.01%** | [−47.20%, −46.16%] | 10.690 → 5.681 |
| `pptx_streaming_create` large | 8+8 × 15 | 190.417 | 171.484 | **−9.89%** | [−12.39%, −7.75%] | 192.953 → 173.531 |
| `pptx_streaming_create` medium | 8+8 × 30 | 6.378 | 5.835 | **−8.57%** | [−8.85%, −8.27%] | 6.416 → 5.978 |
| control: `docx_semantic_one_edit_save` large | 8+8 × 40 | 5.590 | 5.624 | +1.19% | [−0.08%, +1.53%] | 5.697 → 5.734 |
| control: `xlsx_ordinary_save_lifecycle` | 8+8 × 20 | 21.388 | 19.398 | −3.50% | [−9.00%, +10.07%] | 31.464 → 32.068 |

Every process of a leg publishes the same package: before `adf297d1…`,
`62ad295c…`, `cf6965fc…`, `c8ac0ef5…`, `c7b08da6…`, `1f33f8b2…`; after
`e263d219…`, `04e3ea69…`, `edbf573c…`, `7a708509…`, `8658e17d…`, `6cfa2234…`
(the same digests the separate `member_probe` processes print). The XLSX
ordinary-save control publishes `20335c44…` on both legs. The DOCX
semantic-edit control reports no digest; its sink summary is identical on both
legs (37,338 bytes in 51 writes, the same size buckets) and the preservation
probe below shows its path's bytes unchanged.

### Instructions and cycles

User-space counters per timed iteration, by differencing a low- and a
high-sample run of the same binary (`scripts/perf_delta.py`), which cancels
corpus construction and the reopen gates. Two rounds of before, after, after,
before; medians of four measurements per leg and the median paired change
(`counters/summary.json`):

| case | instructions before → after | change | cycles before → after | change |
| --- | --- | ---: | --- | ---: |
| DOCX large | 1,152.91M → 883.41M | −23.39% | 222.69M → 166.71M | −25.59% |
| DOCX medium | 72.94M → 56.06M | −23.15% | 14.46M → 9.76M | −32.00% |
| XLSX large | 2,386.04M → 1,966.87M | −17.57% | 730.15M → 404.16M | −44.28% |
| XLSX medium | 151.36M → 122.37M | −19.15% | 48.83M → 24.82M | −49.17% |
| PPTX large | 2,888.94M → 2,696.40M | −6.99% | not interpretable | — |
| PPTX medium | 98.41M → 93.28M | −5.21% | 29.32M → 23.20M | −19.58% |
| control: semantic edit | 349.06M → 349.09M | −0.02% | 72.07M → 72.14M | +0.02% |
| control: ordinary save | 74.14M → 74.14M | +0.01% | 29.28M → 28.19M | −2.94% |

Each PPTX large process spends about 120 G cycles building and reopening its
16,421-member corpus outside the timer, so its cycle difference is lost in
that noise (paired changes from −253% to +181%); its instruction
measurements agree within about 1% on each leg. XLSX loses 18% of its
instructions but 44–49% of its cycles: level 5's shorter match search skips
the long chain walks of `longest_match`. The controls execute the same instructions to within 0.07%.
The ordinary-save control's wall-clock spread (process p50s from 14.2 to
34.4 ms, moving by round with the host's load) is noise on unchanged work.

### Allocations

Allocator counters per timed iteration from the allocation build
(`alloc/summary.json`; one process per leg and case, medians over the
samples):

| case | allocation calls | allocated bytes | region peak live bytes |
| --- | --- | --- | --- |
| DOCX tiny / medium / large | 105 / 8,233 / 131,113 → +1 each | +16,416 each | +16,414 each |
| XLSX tiny / medium / large | 136 / 8,264 / 131,144 → +1 each | +16,416 each | +16,414 each |
| PPTX tiny / medium / large | 852 / 9,051 / 278,155 → +1 each | +16,416 each | +16,414 each |
| control: ordinary save | 29,845 → 29,845 | unchanged | 14,978,576 → 14,978,574 |

The extra allocation is the 16 KiB stage, made at the first owned Deflate
write and lent to every later member; the other 32 bytes are the reusable
state's new fields. The semantic-edit control reports no allocator counters.

### Sizes per corpus

`member_probe digest` (built against each tree) writes the nine harness
corpora (tiny, medium, large) and lists every member's name, uncompressed
length, SHA-256 of its uncompressed bytes and compressed size
(`probe/member-digest-summary.txt`):

| corpus | members | before archive bytes | after archive bytes | change |
| --- | ---: | ---: | ---: | ---: |
| DOCX tiny | 3 | 1,233 | 1,215 | −1.46% |
| DOCX medium | 3 | 24,571 | 24,553 | −0.07% |
| DOCX large | 3 | 376,848 | 376,054 | −0.21% |
| XLSX tiny | 6 | 3,451 | 3,455 | **+0.12%** |
| XLSX medium | 6 | 167,418 | 163,790 | −2.17% |
| XLSX large | 6 | 2,563,433 | 2,559,706 | −0.15% |
| PPTX tiny | 53 | 36,259 | 35,924 | −0.92% |
| PPTX medium | 549 | 274,398 | 270,370 | −1.47% |
| PPTX large | 16,421 | 7,940,406 | 7,821,547 | −1.50% |

The member lines — 17,050 names, lengths and uncompressed SHA-256s — are
identical on both legs (SHA-256 of all of them `bd7f0916…` on each). The one
growth is the 64-row XLSX corpus's 13,425-byte sheet (1,521 → 1,556 bytes)
that now goes to the codec in one call instead of 64 row writes; its five
small members each lose 6–7 bytes to the dropped sync flush. At level 6 the
same corpus is 3,459 bytes (+0.23%), so the growth is the chunking, not the
level. No change exceeds 1%.

## Error timing and limits

What moves, precisely:

* **Acceptance.** An owned Deflate `write` accepts at most up to the next
  16 KiB member boundary and never fails after accepting a byte; `write_all`
  sees the same total acceptance.
* **When output and its failures surface.** A byte reaches the codec when its
  chunk is complete — at the first write after the chunk fills, or at an
  explicit flush or the finish — and the output that codec call produces is
  written to the sink before that call returns. Before, a byte reached the
  codec in the write that accepted it and its output reached the sink at the
  next write. So compressed output, and therefore a sink failure, an
  owned-entry compressed-limit refusal (`with_compressed_limit`,
  `StreamingArchiveLimits::max_compressed_size`), an output-byte refusal
  (`max_output_bytes`, the DOCX and XLSX writers' `OutputBytes` budget and
  `max_output_bytes`, the PPTX writer's `BudgetedSink`), surfaces up to one
  chunk (16 KiB of uncompressed input) later than before, on a later write or
  at the finish. "Before" already includes the codec's own buffering, which
  on highly compressible input is far longer: the review below measured a
  compressed limit of 7 bytes on 4 MiB of repeated XML refused only after
  4,079,616 input bytes. The writers already mapped a failure at any of those
  calls to the same typed error.
* **What does not move.** The uncompressed entry and total limits
  (`max_entry_size`, `max_total_size`) and every semantic limit of the three
  writers are checked before a write is forwarded, exactly as before. Every
  sink-side limit is still checked before the bytes reach the sink: the
  compressed-limit check refuses a drain whose bytes would exceed the ceiling,
  so no limit is ever exceeded on output. And the decision is exact: an entry
  succeeds if and only if its whole compressed stream fits.

The tests check these: a compressed limit refuses on the write that sends the
chunk's output, with nothing past the ceiling on the sink; across seven
limits from 0 to the stream's size plus one, under random write splits, the
entry finishes exactly when the stream fits and the sink never receives more;
the bounded streaming archive's output and compressed limits at N−1, N and
N+1 succeed exactly at N and N+1, with the refused archive a prefix of the
unlimited one; a failing sink surfaces on the write whose chunk reaches the
codec, which accepts none of its input.

## Correctness evidence beyond the tests

* **Members are unchanged.** `member_probe digest`, built against both trees
  with the same source, lists all 17,050 members of the nine corpora; the
  member lines are identical (SHA-256 `bd7f0916…` on both legs). Only the
  compressed encoding differs. The harness's reopen gates (semantic text,
  paragraph, cell and slide checks against the generators) pass on both legs
  in every measured process.
* **Determinism across writes and processes.** `member_probe split SEED`
  writes the large DOCX corpus handing every run's text to `write_text` whole
  (seed 0), in 1- and 2-byte pieces (seeds 1, 2) and in random 1–23-byte
  pieces (seeds 3, 4), each in its own process, plus the large XLSX and PPTX
  corpora (`probe/split-{before,after}.txt`). After: one digest per format
  across the five processes. Before: the 2-byte splitting gives a different
  DOCX package (`da9a259b…` against `adf297d1…`). The 16 harness processes per
  streaming case agree within each leg.
* **Preservation bytes are unchanged.** The code the preservation writer runs
  (`preserve.rs`: `ReusableDeflateState::new`, `begin_member`,
  `ReusedDeflateEncoder`) keeps level 6 and one codec call per write;
  `begin_member` only adds a level check that is false for a state no owned
  member touched. The XLSX ordinary-save control publishes `20335c44…` on both
  legs, and `member_probe preserve`, run on all 62 DOCX fixtures of the
  workspace (open, replace the first paragraph's text through the public
  semantic edit, publish through the preservation writer), prints the same
  SHA-256 on both legs for the 23 that publish and the same refusal for the 39
  the edit refuses (`probe/preserve-{before,after}.txt`).
* **Mutation checks.** Restoring one codec call per write in the owned path
  fails seven of the new `deflate_reuse` tests and the new DOCX test;
  restoring the pre-finish sync flush fails six. Both were reverted and the
  file compared equal.

## Durable formats

No durable format binds a streaming writer's compressed bytes: packages are
created, not replayed, and the durable patches that bind physical output (for
example PPTX `LPRM`/`LPCP` and the DOCX tail-append proof patch) bind packages
opened and republished through the preservation writer, whose bytes do not
move. A
patch made against a package the old writer created still applies to those
bytes. Nothing is bumped.

## What is not claimed

No claim is registered. The numbers are scoped to the named synthetic corpora,
this host, CPU 12 and the harness lock (flate2 1.1.9 over zlib-rs 0.6.7; the
workspace lock has flate2 1.1.10 over the same zlib-rs). They do not establish:

* the level-5 size effect on content unlike the workspace fixtures; the real
  fixture table is the evidence, and it covers producer XML and media, not
  arbitrary text fed through the streaming writers;
* compression ratios or timings under other zlib backends or zlib-rs versions:
  both the chunk dependence and the level parameters are zlib-rs 0.6.7's;
* RSS, cold cache, other sinks, other producers or other platforms;
* anything about the budget charges, which are record
  [0763](0763-rough-budget-leases.md)'s subject.

## Verification

All gates pass at `bc18e8abdd` except the known base failure listed below
(`results/change-0762/gates.txt`). They are: `cargo fmt --all --check`;
`cargo check --all-targets` of `soapberry-zip` and every workspace package that
depends on it (26, ODF included, iWork excluded; no warnings); warning-denied
Clippy on the library and all targets of `soapberry-zip` and `litchi-docx`;
warning-denied rustdoc of both; `cargo test` of `soapberry-zip` (652),
`litchi-docx` (1,878), `litchi-opc` (918), `litchi-xlsx` (2,036), `litchi-pptx`
(1,167), nine other OOXML and OLE2 dependents and the facade with
`doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (382), plus, as an extra, the eleven ODF
dependents; crate boundaries; the structural claims check (10 claims).
`tools/non_iwork_gate.py verify` fails with "unexpected: litchi-xldm", as it
does at the base (a known failure being fixed separately). The harness is
unchanged; its tests and the coverage validator were not run.

New and replaced tests. `soapberry-zip` `tests/deflate_reuse.rs`:
`staged_model_agrees_with_flate2_encoder_fed_the_same_chunks`,
`owned_deflate_members_follow_the_staged_protocol_across_plans_and_store_members`
(replaces the per-write reference comparison; the plans now include writes
straddling chunk boundaries and an Office-shaped XML member),
`owned_deflate_bytes_do_not_depend_on_write_sizes_or_sink_behaviour` (24
random splittings, half of them Office-sized writes of at most 64 bytes;
short sinks of 1, 3 and 4,097 bytes; sinks interrupting every 2nd, 3rd and 7th
call), `explicit_flushes_cut_chunks_but_keep_the_absolute_boundaries`,
`owned_deflate_accepts_input_up_to_each_chunk_boundary`,
`owned_deflate_reports_a_sink_failure_when_its_chunk_reaches_the_codec`,
`owned_deflate_compressed_limit_refuses_at_the_codec_and_is_never_exceeded`,
`owned_and_borrowed_members_of_one_archive_keep_their_own_protocols` (the
borrowed member after an owned one equals a fresh level-6 encoder with the
sync flush), and
`streaming_output_and_compressed_limits_stay_exact_on_batched_members`.
`litchi-docx` `tests/streaming.rs`:
`package_bytes_do_not_depend_on_how_run_text_is_split` (600 paragraphs, each
run's text written whole and in four piece patterns; identical packages that
reopen to the same paragraphs).

## Review corrections

An independent review (its probes are not part of this packet) confirmed the
change: 60 random write splittings (zero-length, boundary-ending and
multi-chunk writes, raw `write` loops, short and interrupting sinks) and 30
with flushes at fixed offsets gave identical archives; 1,920 preservation and
borrowed-path digests over 320 OOXML fixtures were identical on both builds;
compressed and output limits were exact at N−1, N, N+1 and T−1, T, T+1. It
found two defects in the owned-entry writer, both present at the base but more
likely to matter with batching, and one inaccurate sentence; all three are
fixed in `7b0bd40c02`:

* **A poisoned entry still wrote.** After a sink failure other than
  `Interrupted` poisoned the archive, a later `ZipOwnedEntryWriter::write` or
  `flush` still ran: a Deflate entry pushed the chunk the failed write had
  left staged into a sink that had since recovered, and answered `Ok` (at the
  base it sent the failed write's pending output instead). Only `finish`
  refused, and only after sending that chunk and the final block. `write`,
  `flush` and `finish` now check the poison before any sink I/O.
  `StreamingArchiveEntry`, which the Office writers use, already refused a
  poisoned entry itself. Test:
  `a_poisoned_owned_entry_refuses_later_calls_before_its_sink` (Deflate and
  Store; the sink fails at 100 bytes to 40 KB and is then healed; a later
  write, an empty write, a flush and the finish are all refused, and the sink
  receives nothing more).
* **An interruption in `finish` lost the archive.** `finish` consumes the
  entry, so a sink call interrupted with `Interrupted` inside it ended the
  entry and its archive, and batching moves up to one chunk's output into
  `finish`. `finish` now retries an interrupted sink call where it stopped, as
  `write_all` does; every codec call still follows a complete drain, so none
  is repeated and the member's bytes do not change. Tests:
  `owned_entry_finish_retries_interrupted_sink_calls_without_changing_bytes`
  (the test corpus finished into sinks interrupting every 2nd, 3rd or 5th
  write or flush equals the reference archive), and the split test's
  interrupting sinks now stay armed through each entry's `finish`. The
  archive's own `finish` ends with a sink flush that does not retry an
  interruption; that is unchanged.
* **Rustdoc.** `ZipOwnedEntryWriter` and `StreamingArchiveEntry` said a
  refusal surfaces "up to one 16 KiB chunk of input later than the write that
  caused it". It surfaces after the codec's own buffering plus up to one
  staged chunk, as the section on error timing says; both now say so.

Removing either poison check or either retry fails a test. On success paths
the fix adds one poisoned-flag test per owned-entry `write` and `flush`; the
timings above were not re-measured. The gates after the fix are listed in
`results/change-0762/gates.txt` and in
[0763](0763-rough-budget-leases.md#review-corrections).

## Cleanup

Binary digests are in `binaries.txt`. After the evidence was copied, the
target directories were removed: `targets/0762` (221 GB of full-debug test
builds, removed early when the shared disk filled to 0 bytes during the gate
runs, then rebuilt at 1.4 GB without debug information for the six ODF test
reruns), `targets/0762-before`, and the three probe targets. The probes' full
outputs and perf data are summarized in the packet. The detached before
worktree `0762-before-src` is re-used, checked out at `bc18e8abdd`, as the
before leg of [0763](0763-rough-budget-leases.md), which removes it with the
remaining target and scratch directories (`cleanup.json`).
