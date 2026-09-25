# 0769 — CFB open and every reader agree on the partial last mini sector: a stream the root size covers now reads, a stream it cuts is refused at open, and 0767's 52 read errors after open are gone

Status: retained, implemented in `litchi-cfb`. `performance_claim: none`.
This is a correctness change. Its timings, instruction and cycle counts,
Callgrind counts and allocation counts are evidence of what it costs, not
registered claims.

OLE2 and OOXML remain the active priority. ODF work stays deferred until that
goal completes; iWork is excluded.

Base `97e558cef4` (record [0767](0767-cfb-reparse-linear.md)'s head, under
review); branch `perf/0769-cfb-mini-sector-open-read-agreement`.

- `97bf5f8f82`: record 0767's review follow-ups (see
  [below](#record-0767s-review-follow-ups)).
- `d13c52a45f`: the agreement. A5 and the three whole-stream readers bound
  the bytes a mini stream takes. Tests, and the vendored real-producer
  fixture.
- `47d19bd609`: record 0767's fault lane as a regression test.
- `ebf20bba87`: the refinement. A5's ownership loop is the base's again, and
  the partial sector is checked once per stream. This is the measured
  candidate.

## Result

Record 0767's fault lane found that open-time validation (A5) and
`open_stream` disagree on the same bytes. A5 admitted a mini stream in the
partial last mini sector of a root mini stream whose size is not a multiple
of 64. `open_stream` then refused it with "Mini sector out of bounds". This
caused all 52 of that lane's read errors after a successful open.

MS-CFB does not require the root size to be a multiple of 64, and a real
producer writes it unaligned (below). The rule is now the same at open and on
every read: a mini stream is admitted exactly when every byte it takes from
its mini sectors lies inside the root size.

- **A stream the root size covers now reads through every reader.**
- **A stream that needs a byte past the root size is refused when the file
  is opened.** It gets the typed error A5 already gave mini storage outside
  the root mini stream, `CorruptedFile("Mini stream references storage
  outside the root mini stream")`. Before, it was refused only when read.

| Lane (base → candidate) | Admitted by the open, then a read refused | Readers disagree on a stream | Other verdict changes |
|---|---:|---:|---:|
| Record 0767's fault lane: 30,400 byte-level faults + 104 clean files, same seed | 52 → **0** | — | 0 |
| Root-size sweep: every root size within 128 bytes of each file's own, 105 files, 21,741 copies | 3,906 → **0** | 1,353 → **0** | 0 |
| Census: all 1,830 compound files under `test-data/` and the reference corpora, plus the normalized real-producer file, six readers per stream | 1 → **0** | 1 → **0** | 0 |

- **The fault lane's 52 cases now read.** Every stream the root size covers
  reads through every reader. The base arm reproduces record 0767's recorded
  outputs byte for byte.
- **The sweep moves 3,906 copies.**
  - 1,353 were admitted with a stream that some readers returned and others
    refused. They now read through every reader.
  - 2,553 were admitted with a stream no reader could return. They are now
    refused at open.
  - All 21,741 candidate verdicts match an independent byte-bound oracle.
- **No real file changes verdict.** Of 1,830 compound files, only 3 have a
  root size that is not a multiple of 64, and litchi refuses all three at
  open for other reasons.
- **The cost is within noise.**
  - An open runs about 3 more instructions per stream (+0.04% at 10,000
    mini streams), and 48 to 84 more per real fixture.
  - Reading every mini stream runs 1.5% fewer instructions.
  - Allocations are identical in all 17 probe cases.
  - No case's median time moved more than +3.5%. Every move is explained by
    code placement: the candidate executes as many instructions or fewer.

## The disagreement

The root entry's stream size is the size of the mini stream (MS-CFB 2.6.1).
Mini sector *n* covers bytes 64*n*..64*n*+64 of it. When the root size is not
a multiple of 64, the last mini sector is partial. Before this record, the
layers of `litchi-cfb` bounded that sector three different ways:

| Layer | Bound on a mini sector | A stream whose bytes end inside the partial sector |
|---|---|---|
| A5 (`validate_stream_allocations`, every `OleFile::open` and `SharedOleFile` open) | the sector index is below ⌈root size / 64⌉ | admitted, even when the stream needs bytes past the root size |
| `OleFile::open_stream`, the writers' in-place readback (`StreamComparer`), `SharedOleFile`'s mini-stream cache | the whole 64-byte sector lies inside the root size | refused ("Mini sector out of bounds"), even when every byte it needs lies inside |
| `OleFile::read_stream_range`, `SharedOleFile`'s bounded direct read, ranges and cursor, the same-length overlay, the splice | exactly the bytes read lie inside the root size | read when its bytes lie inside, refused otherwise |

One consequence went beyond open against read. **One `SharedOleFile`
could answer both ways for the same stream.** Its `open_stream` takes the
bounded direct path for an eligible small target and the cache after a
different target, and cache takeover is permanent. So the same stream of the
same reader was returned or refused depending on which streams were read
before it. The sweep shows readers disagreeing in 1,353 base copies.

## What MS-CFB says, and what producers write

- **MS-CFB 2.6.1:** "For a root storage object, this field contains the size
  of the mini stream." No section requires that size to be a multiple of 64;
  2.4 says only that the mini stream "is divided into smaller, equal-length
  sectors".
- **MS-CFB 2.7:** a sector chain "MUST be greater than or equal to the stream
  size", and "the unused portion of the last sector of a stream object's
  user-defined data SHOULD be filled with zeroes". A stream's bytes are
  therefore the first `size` bytes of its chain's sectors: all of each sector
  but the last, and the rest of the stream in its last. The unused tail of
  the last sector holds none of them.
- **The example in MS-CFB 3.3–3.5:** one 544-byte stream (8.5 mini sectors)
  and a 576-byte root, nine whole mini sectors. So a root size can equal the
  allocated mini sectors, but nothing requires it.

Independent census (`scripts/census.py`, its own CFB reader, no litchi code)
of every compound file available here: 98 under `test-data/ole`, 116
elsewhere under `test-data`, and 1,616 in the gitignored POI, LibreOffice and
Open-XML-SDK reference corpora.

- **The independent reader parses 1,801 files, and 1,798 of them have a
  root size that is a multiple of 64.** It cannot parse 29 files, and cannot
  walk some mini chain of 37 others; litchi refuses all 66 at open.
- **Three do not:**
  - **POI's `BlockSize512.zvi` and `BlockSize4096.zvi`** are Zeiss
    AxioVision images. Their root size is 11,564 bytes: 180 mini sectors
    and 44 bytes. The last mini stream, `\x05DocumentSummaryInformation`
    (172 bytes, mini sectors 178, 179 and 180), ends exactly at the root
    size, 44 bytes into the partial sector. **A real producer writes the
    root size as the end of its last stream's bytes.**
  - **One POI fuzzer case**, with an invalid mini sector shift.
- **litchi refuses all three at open for other reasons:**
  - `BlockSize512.zvi`: the producer writes `FREESECT` into its storages'
    start sector fields, which the directory validation refuses ("invalid
    CFB storage fields at SID 2").
  - `BlockSize4096.zvi`: a version-3 header with 4096-byte sectors.
  - The fuzzer case: its mini sector shift is 14.

  So no file litchi opens today changes verdict. With only those storage
  fields normalized (`scripts/normalize_zvi.py`), `BlockSize512.zvi` opens.
  In the base its summary stream is then refused. In the candidate it reads
  with SHA-256 `6f74c82e…c0efdbc9`, the digest the independent reader
  computes.

Two reference readers, checked in the local sources:

- **LibreOffice** (`sot/source/sdstor/stgstrms.cxx`): `StgSmallStrm::Read`
  reads each mini page through the root stream. `StgDataStrm::Read` clamps
  that read to the root size, so it returns every byte a stream needs that
  lies inside the root size, and no more.
- **Apache POI** (`POIFSMiniStore.getBlockAt`) slices mini blocks out of the
  root's big-block chain without consulting the root size, and writes the
  root size as the mini block count × 64 (`RootProperty.setSize`).

## The decision

A mini stream is admitted exactly when every byte it takes from its mini
sectors lies inside the root size. The unused tail of its last sector may
extend past the root size. That can only happen at the partial mini sector,
and only for the sector that ends the stream.

- **Accept at read** what the spec allows and a real producer writes (0652
  trade-off 3). Every byte such a stream needs lies inside the root size.
  A5's exact root chain check keeps each of those bytes in the root's
  allocated sectors too.
- **Refuse at open** a stream that needs a byte past the root size (ADR
  0006: "fatal safety failures stop opening immediately"). The refusal has
  the same type (`CorruptedFile`) and the same message as any other mini
  storage outside the root mini stream. A5 already refuses the whole file for
  every other mini allocation fault, so this is where the refusal belongs.

**The rejected alternative** refused any stream in the partial sector at
open, treating the mini stream as ⌊root size / 64⌋ sectors. It is simpler,
but it refuses files that MS-CFB allows and that a producer writes, and
litchi's own range readers and overlay already follow the byte bound. (The
splice checks only the first byte of each per-mini-sector span, at
`splice.rs:489`; that is harmless because a span stays inside one 64-byte mini
sector within the root chain's last sector, and A5 now bounds every byte —
review correction.)

## What was changed

`crates/litchi-cfb/src/file.rs`:

- **`MiniStreamEnd`** holds the root size as whole mini sectors plus the
  partial sector's bytes. `has_partial_sector` says whether there is a
  partial sector. `admits_chain_end` says whether a stream's chain uses the
  partial sector only as its last sector, for no more bytes than the root
  size covers.
- **A5.** The per-sector ownership loop is the base's, comparing with the
  capacity rounded up to whole mini sectors. After a stream's chain is
  claimed, and only when the root size has a partial sector, the stream is
  held to `admits_chain_end`. The check runs once per stream and is out of
  line (`#[cold]`).
- **`read_minifat_chain`** (the `open_stream` mini path) and
  `StreamComparer::minifat_stream_equals` bound the bytes they copy from each
  sector, not the whole sector. They stop at the stream's last byte, so
  sectors after it are not read.

`crates/litchi-cfb/src/shared.rs`:

- **`SharedOleFile::read_minifat_stream`** (the mini-stream cache) bounds the
  bytes it copies in the same way.

The "Mini sector out of bounds" checks stay, with their strings, as defence
in depth. On a validated file they can no longer fire.

Tests and fixture:

- `crates/litchi-cfb/src/mini_stream_end_tests.rs`, new;
- record 0767's A5 oracle in `chain_work_tests.rs`, which follows the new
  bound;
- `test-data/poi/test-data/poifs/BlockSize512.zvi`, vendored unmodified
  (SHA-256 `4454d647…a395c49c`, 51,712 bytes), next to the other POI
  fixtures under `test-data/poi` (Apache License 2.0, `NOTICE` there).

There is no public API change, no new dependency, no `unsafe`, no new or
changed limit, and no change to planning, emission or any published byte.
The writers produce root sizes that are multiples of 64, as before. A
reused-layout save of an unaligned source still declines
(`SourceMiniStreamUnaligned`) and serializes from scratch.

## What changes, and what cannot

**Verdicts change only for files where some mini stream uses the partial
last mini sector of a root whose size is not a multiple of 64:**

| Such a stream | Base | Candidate |
|---|---|---|
| every byte it takes lies inside the root size | admitted; whole-stream reads refuse it, range reads and the shared direct read return it | admitted; every reader returns it |
| it needs a byte past the root size | admitted; every reader refuses it | refused at open (`CorruptedFile`, "outside the root mini stream") |

- **Such a file is now refused as a whole.** A caller could open it before
  and read its other streams, but can no longer, as for every other mini
  allocation fault A5 refuses.
- **Nothing else changes.**
  - For a root size that is a multiple of 64 there is no partial sector.
    `has_partial_sector` is false, A5 runs exactly the base's checks, and
    each reader's bound over a full sector equals the old whole-sector
    bound.
  - A reader stops at the stream's last byte, where it used to bound one
    sector past it. A validated chain has no sector past it: A5 walks every
    mini chain to exactly ⌈size / 64⌉ sectors and `ENDOFCHAIN`. So this
    differs only on tables faulted after a validating open, which only
    crate tests build.
  - **Error order among several faults in one stream.** A5 claims the
    stream's whole chain before its partial-sector check. So a sector
    claimed twice, or past the rounded-up capacity, is reported before a
    partial-sector refusal of the same stream. That order is new, because
    the base had no partial-sector refusal.
- **The other layers already followed the byte bound:** the range readers,
  the shared direct read and cursor, the same-length overlay
  (`derive_selection_spans`) and the splice. On a validated file their
  checks can no longer fire either. A file they used to accept is now
  admitted, and one they refused is now refused earlier.
- **The in-place readback reaches `open_stream`'s verdict,** because both
  changed identically (`stream_compare_tests.rs` holds them together over
  every fixture).

## Record 0767's review follow-ups

The review of record 0767 (merge, with two nits) asked for these in the same
file. They are a separate commit, `97bf5f8f82`:

- **The MiniFAT read now moves the reader's chain scratch out and puts it
  back, as the FAT read does.** The root mini stream's load, itself a FAT
  read, runs first and puts its own buffers back. So a panic mid-read on
  either path leaves the reader an empty, all-clear scratch.
  `EndChainScratch::collect` and `SectorChainScratch::collect_exact` now
  `debug_assert!` that their visited map is all-clear on entry.
- **A reader frees its chain buffer after a read once the buffer's capacity
  exceeds `RETAINED_CHAIN_SCRATCH_BYTES`** (1 MiB: a chain of more than
  131,072 sectors, a stream of more than 64 MiB even at 512-byte sectors).
  Smaller buffers are kept, so the common path is unchanged. The allocation
  lane is identical in every case.
  - Before, the buffer kept one `u32` per sector of the longest chain ever
    read, about 8 MiB after one 1 GiB stream of 512-byte sectors.
  - The visited map is kept whatever its size. It holds one bit per entry of
    a table whose `u32` entries the reader already holds for its lifetime (a
    thirty-second of that memory). Freeing it would bring back the per-read
    table-sized clear that record 0767 removed. **This departs from the
    review's wording, which named both buffers; the reason is stated at the
    constant.**
- **Tests:**
  - an oversized buffer is freed after FAT and MiniFAT reads and after a
    failed read;
  - a buffer exactly at the threshold is kept;
  - the map is kept and all-clear;
  - mini-only read sequences keep one scratch.
- **Mutation checks:** removing either release, or the MiniFAT put-back,
  fails a test. A restoration that never clears trips the new assertion in
  166 tests.

## Tests

`mini_stream_end_tests.rs` holds every expectation to an independent
byte-bound oracle. The oracle is the least root size that holds every byte
every mini stream takes, computed from sizes and chains, not from the
validation code. Every admitted copy must return the expected bytes through
eight readers:

- `OleFile::open_stream`;
- `OleFile::read_stream_range`, whole and its last byte;
- `StreamComparer`;
- a fresh `SharedOleFile`'s `open_stream`, the bounded direct path when
  eligible;
- the shared mini-stream cache;
- the shared range reader;
- the shared cursor.

Every refused copy must be refused by both opens with the same message.

- **`root_sizes_around_a_partial_last_mini_sector_follow_the_byte_bound`.**
  - v3 and v4 files, whose last mini stream is 1 to 4,095 bytes.
  - Root sizes 64*n*−1, 64*n* and 64*n*+1 for the last two mini sectors,
    and the byte bound minus one, at, and plus one.
  - 176 copies: 46 newly admitted and 58 newly refused.
  - Every newly admitted copy is also republished over its adopted source.
    The reused layout declines (`SourceMiniStreamUnaligned`), and the
    output reads back identically with an aligned root.
- **`a_partial_mini_sector_inside_a_chain_is_refused_at_open`:** a chain
  permuted so that its first sector, all 64 bytes of it needed, is the
  partial one.
- **`seeded_mini_layouts_follow_the_byte_bound`:** 160 seeded files with
  their mini sectors shuffled, and random root sizes. 640 copies, which
  reach every outcome.
- **`a_real_producer_stream_ending_in_the_partial_mini_sector_reads`:**
  - pins the premise of `BlockSize512.zvi` (root 11,564; the summary stream
    in mini sectors 178, 179 and 180);
  - pins its unrelated refusal as it is;
  - normalizes the storage fields in memory, then pins the stream's
    independently computed SHA-256 and the refusal one byte short.
- **`every_admitted_fixture_stream_reads_through_every_reader`:** the census
  of all 215 compound files under `test-data`. 211 are admitted, with 1,303
  streams through every reader; the 4 refusals are the same from both
  opens.
- **`root_sizes_around_each_fixture_mini_stream_end_follow_the_byte_bound`:**
  the root-size cases above over the 61 `test-data/ole` fixtures that have
  mini streams. 549 copies: 180 newly admitted and 173 newly refused.
- **`every_stream_of_an_admitted_faulted_file_reads_through_every_reader`**
  (`47d19bd609`) is record 0767's fault lane as a regression test.
  - Seeded faults hit FAT and MiniFAT links, stream and root entries, and
    the root size.
  - The inputs are written files and every ninth fixture, including a file
    whose mini streams are whole sectors, so a root one byte short cuts a
    stream.
  - 900 copies: 258 admitted and read through every reader, 642 refused
    alike.

**Mutation checks** (on the final code and tests):

- **Disabling the per-stream check**, which restores the base's A5 rule,
  fails six of the seven new tests. The census test passes on the base as
  well, because no fixture litchi opens has the shape. The fault test
  catches this only because its whole-sector inputs were added for the
  purpose.
- **Restoring the whole-sector bound** in `open_stream`, `StreamComparer` or
  the shared cache, one at a time, fails five tests each.

## Cross-build lanes

Both arms' binaries ran on core 30 (`lanes/` in the packet).

**Record 0767's fault lane**, same 104 inputs, case counts and seed. The
base arm's outputs are byte-identical to record 0767's recorded ones.

| Arm | Cases | Opened | Refused | Reads returned | Reads refused | Opened, then a read refused |
|---|---:|---:|---:|---:|---:|---:|
| base | 30,504 | 2,692 | 27,812 | 53,257 | 52 | 52 |
| candidate | 30,504 | 2,692 | 27,812 | 53,309 | 0 | **0** |

- **All 52 differences are cases whose streams now read.** A root
  shortened to 64*n*−1 cut a stream whose tail needs at most 63 bytes.
- **No case needed a refusal at open**, and no other case changed.

**The root-size sweep.** For each of those 104 inputs and the normalized
Zeiss file, every root size within 128 bytes of the file's own (from 1). The
readers are those of the census below, without a fresh shared reader per
stream.

| Arm | Admitted | Refused | Admitted, then a read refused | Readers disagree |
|---|---:|---:|---:|---:|
| base | 11,247 | 10,494 | 3,906 | 1,353 |
| candidate | 8,694 | 13,047 | **0** | **0** |

- **Every difference is one of the two classes above:** 1,353 copies now
  read, and 2,553 are now refused at open.
- **The candidate matches the independent oracle** (`census.py`'s reader,
  each file's own metadata) on all 21,741 copies:
  - 8,694 admitted;
  - 6,254 outside the root mini stream;
  - 6,793 refused because the root chain's length changes.

**The census.** 1,830 compound files and the normalized Zeiss file, both
opens and six readers per stream.

- **Base:** 1,647 admitted, 23,475 streams. One file has a stream some
  readers refuse: the normalized Zeiss file's summary stream.
- **Candidate:** the same verdicts, and every reader returns every stream.
- **That file is the only difference.**

## Method

- **Arms.** Base `97e558cef4` (a detached checkout) and candidate
  `ebf20bba87`; the first version, `d13c52a45f`, has its own matrix and
  counters.
  - All were built with identical cargo commands, flags and features under
    rustc 1.95.0 (`binaries.json` has the commands and every SHA-256).
  - The probe manifests differ only in their dependency tree, and
    `--remap-path-prefix` maps both trees to `/litchi`.
  - A temporary test mutation of `file.rs` overlapped the first candidate
    harness build. The measured harness is a second build with the identical
    command, which recompiled `litchi-cfb` and its dependents from the
    committed source. The probes were built only from committed source.
- **The probe** is record 0767's (`probe/`), with a `census` and a
  `root-sweep` mode added. The twelve generated inputs are record 0767's,
  regenerated byte for byte.
- **Timing.**
  - 16 rounds of 27 cases, ABBA across rounds with rotated case order.
    Each binary starts through a symlink 8 bytes longer each round (record
    0745's layout sampling). Every process was pinned with
    `taskset -c 30`. 608 processes ran and none failed.
  - The cases:
    - `OleFile::open` of the v3 and v4 1,000- and 10,000-mini-stream files
      and a 10,000-regular-stream control;
    - reading every stream of the same files;
    - `SharedOleFile::open_owned` and the common editor's open-and-capture;
    - the Reuse write of 45543.ppt (its readback runs `StreamComparer`);
    - opens of three real fixtures, and a Rewrite-write control;
    - the harness selectors `cfb_open` and `cfb_list_streams` (four shapes)
      and `doc_semantic_open` (tiny, large).
  - Statistics are record 0767's: a comparison is the median of per-round
    paired changes, with a 10,000-resample bootstrap interval (seed 769).
- **Counters.** Per-owner `instructions:u` and `cycles:u` come from two
  processes at three argv[0] layouts. Harness rows include each iteration's
  untimed setup.
- **Callgrind.** The harness's `doc_semantic_open/large`, 3 samples, whole
  program, in all three builds.
- **Allocation.** One counting-allocator process per probe case and arm.

## Results

`summary.md` has every row.

| Case | Base p50 | Cand p50 | Δ p50 [95% CI] | Instructions per owner, base → cand |
|---|---:|---:|---:|---:|
| `OleFile::open`, v3, 10,000 mini streams | 3,138.4 µs | 3,133.8 µs | −0.14% [−0.97, +0.31] | 72,852,954 → 72,883,029 (+0.04%) |
| `OleFile::open`, v4, 10,000 mini streams | 3,039.7 µs | 3,030.6 µs | −0.55% [−1.18, +0.26] | 70,551,710 → 70,581,800 (+0.04%) |
| `OleFile::open`, v3, 1,000 mini streams | 359.3 µs | 357.0 µs | −0.89% [−1.71, −0.13] | 7,296,153 → 7,299,210 (+0.04%) |
| `OleFile::open`, v3, 10,000 regular streams (control) | 1,878.1 µs | 1,891.7 µs | +0.53% [+0.28, +2.39] | 38,708,571 → 38,718,633 (+0.03%) |
| Read every stream, v3, 10,000 mini streams | 4,665.3 µs | 4,605.3 µs | +0.36% [−2.99, +2.13] | 146,346,850 → 144,106,881 (−1.53%) |
| Read every stream, v4, 10,000 mini streams | 5,120.5 µs | 5,055.8 µs | +1.28% [−0.57, +2.29] | 142,176,105 → 139,936,127 (−1.57%) |
| Read every stream, v3, 1,000 mini streams | 338.6 µs | 349.0 µs | +3.29% [+0.99, +4.20] | 14,267,530 → 14,043,562 (−1.57%) |
| Read every stream, v3, 10,000 regular streams | 5,235.0 µs | 5,160.6 µs | −0.21% [−1.53, +1.32] | 87,720,254 → 87,820,384 (+0.11%) |
| `SharedOleFile::open_owned`, v3, 10,000 mini streams | 2,690.0 µs | 2,576.3 µs | −4.44% [−4.84, −4.00] | 57,875,471 → 57,895,334 (+0.03%) |
| Common-editor capture, 3,000 mini streams | 2,817.5 µs | 2,811.8 µs | −0.02% [−0.90, +0.45] | 49,412,494 → 48,665,558 (−1.51%) |
| Reuse `write_to`, 45543.ppt | 23.1 µs | 22.7 µs | −0.97% [−2.74, +1.69] | 378,322 → 378,325 (+0.00%) |
| `OleFile::open`, 45543.ppt | 3.26 µs | 3.38 µs | +3.53% [+2.76, +3.99] | 74,804 → 74,851 (+48) |
| `OleFile::open`, FloatingPictures.doc | 5.37 µs | 5.39 µs | +0.19% [−0.71, +1.69] | 114,139 → 114,211 (+72) |
| `OleFile::open`, WithCheckBoxes.xls | 4.49 µs | 4.51 µs | +0.45% [−0.45, +1.22] | 89,519 → 89,604 (+84) |
| Rewrite `write_to`, 45543.ppt (control) | 8.74 µs | 8.77 µs | +1.10% [−1.41, +3.73] | 94,535 → 94,534 |
| harness `cfb_open` tiny / many-small / few-large / wide-root | 1.52 / 76.2 / 72.6 / 483.8 µs | 1.56 / 77.2 / 69.9 / 487.6 µs | +2.63% / +1.29% / −3.64% / +0.71% | +0.14% / +0.04% / +0.00% / +0.05% |
| harness `doc_semantic_open` tiny / large | 5.95 / 284.4 µs | 5.91 / 290.5 µs | −0.51% / +2.42% [+1.67, +3.08] | −0.65% / −0.63% (per owner, with setup) |

- **Opens.**
  - About 3 more instructions per stream: the per-stream
    `has_partial_sector` test.
  - The real fixtures' +48 to +84 instructions are that test and the
    `MiniStreamEnd` setup.
- **Reads of mini streams** run about 224 fewer instructions per stream. The
  copy loop now bounds each sector once, by the bytes it copies. Read-all
  owners include their untimed open, as in record 0767, so the reads alone
  save about 3 more per stream.
- **Reads of regular streams** run 10 more instructions per read: the
  chain-buffer release check of the follow-ups.
- **Allocation** per owner is identical in bytes, calls, peak live bytes and
  retained bytes in all 17 probe cases. Every owner of a process was
  identical, and every probe output matched across the arms.
- **Callgrind** of `doc_semantic_open/large` (3 samples):
  - the DOC parse (`Document::from_ole_with_options`, inclusive) runs
    39,524,786 Ir in the base and 39,137,906 in the candidate (−0.98%);
  - every `litchi_cfb` function is within 0.8% (`validate_stream_allocations`
    176,213 → 177,513; `read_stream_from_fat` 3,304,983 → 3,305,919);
  - the rest of the difference is in `litchi-doc` code with identical
    source, inlined differently because the generic `litchi-cfb` bodies
    instantiated in its codegen units changed (`TextExtractor::new`
    −593,136 Ir, `ChpBinTable::parse` +62,289).

## The first version and why it was refined

The first version (`d13c52a45f`) put the bound inside A5's per-sector
ownership loop: a comparison with the whole mini sectors, then an
out-of-line check of the partial one. Its own 16-round matrix and counters
(`matrix-v1/`, `counters-v1/`) measured the correctness change as retained,
at a cost:

| Case | First version, instructions Δ | Retained, instructions Δ |
|---|---:|---:|
| `open-v3-mini-10000` / `open-v4-mini-10000` | +0.48% / +0.50% | +0.04% / +0.04% |
| `open-45543` / `open-floating` / `open-checkboxes` | +54 / +154 / +271 | +48 / +72 / +84 |
| harness `cfb_open/many-small` | +0.25% | +0.04% |

- **The loop ran about one more instruction per mini sector.** The out-of-line
  call kept the chain and size live across the loop.
- **The retained version keeps the base's loop** and checks the partial
  sector once per stream, only for an unaligned root.
- **The first version's `doc_semantic_open/large` moved +14.09% [+13.17,
  +14.74]**, with 45 per-round flags, while its per-owner instructions fell
  0.65%. Callgrind puts the two candidates' DOC parse at 39,137,904 and
  39,137,906 Ir, and both fall 0.98% below the base.
  - The retained build moves the same case +2.42%.
  - A +14% to +2% swing between two builds whose DOC path runs the same
    instructions is code placement, which record 0746 measured at up to
    ±40% on this host.

## Regression flags (every change above +5%)

No case's median change is above +5%. The retained matrix has 75 per-round
flags (a round's paired p50, mean or p95 above +5%); `summary.md` lists each
case's count and largest.

- **`ctl-write-rewrite-45543`, 10 flags.** This is the control whose timed
  region runs no reader code. They show the noise floor of an 8.7 µs owner:
  its median is +1.10% [−1.41, +3.73] with equal instructions.
- **`open-v3-regular-10000`, 10 flags** (p95 in scattered rounds). Median
  +0.53%, instructions +0.03%.
- **Three to eight flags each:**
  - `cfb_list_streams/few-large` (8) and `cfb_list_streams/tiny` (5): 10 ns
    steps on 140 to 160 ns calls this change does not touch, with equal
    instructions, as in record 0767;
  - `cfb_open/wide-root` (6);
  - `cfb_open/tiny` (5): median +2.63%, +40 ns, instructions +0.14%;
  - `openread-45543` (5);
  - `editor-open-v3-mini-3000` (4);
  - `readall-v3-mini-1000` (4): median +3.29% with 1.57% fewer
    instructions;
  - `doc_semantic_open/large` (3): median +2.42% with fewer instructions and
    a 0.98% smaller Callgrind DOC parse;
  - `shared-open-v3-mini-10000` (3): median −4.44%.
- **One or two flags each:** the rest (`summary.md`).

The largest medians in each direction come with instruction counts that
cannot explain them, so they are code placement:

- `open-45543` +3.53% (+120 ns, +48 instructions);
- `readall-v3-mini-1000` +3.29%;
- `shared-open-v3-mini-10000` −4.44% (+0.03% instructions);
- `cfb_open/few-large` −3.64% (equal instructions).

The candidate harness's `.text` is 34 KB larger and the probe's 6 KB, from
the reader and validation bodies instantiated for each reader type.

## What is not claimed

- No registered performance change: `performance_claim: none`. The timing
  lanes show the change costs nothing measurable, not that it speeds
  anything up. The 1.5% fewer instructions of mini-stream reads are not
  claimed as a speedup either; their timings are within noise.
- **No claim about real files' verdicts beyond this corpus.** The census
  found no file litchi opens whose verdict changes. The two real files with
  this shape are refused for other reasons.
- No claim that other readers treat a stream that needs a byte past the
  root size as litchi now does. LibreOffice returns a short read there, and
  POI reads past the root size.
- **Scope:** warm, in-memory, serial, on one host, with no cold-I/O,
  concurrency, RSS or cross-platform result. Peak live bytes are allocator
  ownership, not RSS.

## Found, not changed

- **`FREESECT` in storage start sector fields.** The Zeiss AxioVision writer
  puts `FREESECT` in every storage entry's start sector. The directory
  validation accepts only 0 and `ENDOFCHAIN` (itself a tolerance for
  deployed writers) and refuses the file. MS-CFB 2.6.1 requires zero, so
  `FREESECT` is a deviation. Admitting it for storages, whose start field
  addresses nothing, would open `BlockSize512.zvi` unmodified. That is a
  directory-validation question for its own record.
- **A version-3 header with 4096-byte sectors** (`BlockSize4096.zvi`) is
  refused by the header check. MS-CFB 2.2 ties the sector shift to the major
  version, and POI reads the file anyway. Not changed.
- **The cursor reader accepts a mini or regular chain longer than its
  stream**, and the shared reader refuses one. This is reachable only on
  tables faulted after a validating open, because A5 refuses such chains at
  open. Not changed.
- **Placement.** `doc_semantic_open/large` moved +14.1% and +2.4% in two
  candidates that run the same DOC instructions. Changes to the generic
  `litchi-cfb` bodies move `litchi-doc` codegen, and that timing case
  follows the layout more than the work.

## Authority and constraints

- **MS-CFB 2.4, 2.6.1, 2.7 and 3.3–3.5** (local copy under
  `3rdparty/specs/[MS-CFB]`) define the rule.
- **ADR 0006:**
  - the normative specifications define canonical semantics, and readers
    preserve real-world quirks;
  - fatal safety failures stop opening immediately;
  - validation never mutates, and nothing is repaired: the readers return
    the producer's bytes, and a refusal stays a typed `CorruptedFile`.
- **ADR 0005:** no ambient behaviour, threads or I/O. The follow-ups bound a
  reader's retained chain buffer.
- **ADR 0003:** no published byte, patch or inverse changes.
- **ADR 0026:** directory validation stays in `litchi-cfb`, unchanged.
- **GOAL, "LEGACY CFB-SPECIFIC WORK":** every validation, ownership, cycle,
  overlap, FAT, MiniFAT, directory and truncation check is preserved. One
  check is strengthened: a mini stream needing bytes past the root size is
  now refused at open.
- **Change [0652](0652-owner-decisions-for-the-third-wave.md)'s standing
  trade-offs:**
  - 1: the change of verdicts is documented here. It needs no API change.
  - 2: correctness first. The layers agree, and the refusal moves to the
    earliest point.
  - 3: the benign common path is unchanged. A5's loop is the base's, and a
    real producer's shape reads. The malicious minority is refused at open.
- **Change [0758](0758-owner-decisions-2026-09-24.md)** decides nothing this
  record uses. No ODF or iWork crate is touched.

## Verification

`gates.txt` in the packet lists every command and exit code. They ran one
after another after all measurement, on `ebf20bba87`, with
`CARGO_BUILD_JOBS=6`, `CARGO_INCREMENTAL=0`, `RUST_TEST_THREADS=6` and
`TMPDIR` under the scratch directory.

- `cargo fmt --all --check`: pass.
- `cargo check --all-targets --locked --offline`: pass, for three sets:
  - `litchi-cfb` and its eight direct dependents;
  - the OOXML crates that reach it through `litchi-crypto`;
  - the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`.
- Clippy `-D warnings` on the `litchi-cfb` library and all its targets:
  pass.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps`, with and
  without `--document-private-items`: pass.
- Tests:

  | Crate | Passed | Failed | Ignored |
  |---|---:|---:|---:|
  | `litchi-cfb` | 510 | 0 | 1 |
  | `litchi-ole-common` | 255 | 0 | 1 |
  | `litchi-doc` | 1,293 | 0 | 13 |
  | `litchi-ppt` | 1,249 | 0 | 11 |
  | `litchi-xls` | 1,611 | 0 | 1 |
  | `litchi-vba` | 47 | 0 | 0 |
  | `litchi-ograph` | 70 | 0 | 0 |
  | `litchi-crypto` | 43 | 0 | 0 |
  | `litchi-sign` | 25 | 0 | 0 |
  | Facade (`litchi`, all OLE2/OOXML features) | 382 | 0 | 7 |

  `litchi-cfb` has record 0767's 501 tests plus the two follow-up tests and
  the seven new ones. The ignored tests are pre-existing.
- Crate boundaries: pass (65 packages, 244 declarations, 11 existing debt
  items).
- Structural perf-claims check: pass (10 claims).
- `non_iwork_gate verify`: fails with "workspace package inventory mismatch
  (unexpected: litchi-xldm)". This is the known failure at the wave's base,
  which the wave briefing lists and another branch is fixing.

The harness is unchanged, so its tests and the coverage validator were not
run. The briefing's other known failures (two harness tests, the facade's
`clippy::unit_arg` lints, iWork crates under workspace-wide `--all-targets`)
are outside these commands. No shared log, claim registry or coverage index
is edited; `log-sections.md` in the packet has the paste-ready sections.

## Cleanup

`cleanup.json` records every removal.

- **Removed after committing this record and packet:**
  - `targets/0769` and `targets/0769-before`: the gate build, and the probe
    and harness builds of every arm;
  - the `0769-before-src` checkout, with `git worktree remove --force`;
  - `scratch/0769`:
    - staged binaries (hashes in `binaries.json`);
    - generated inputs (regenerable, digests in `inputs.json`);
    - argv0 symlinks, and raw lane and matrix outputs (bundled in the
      packet);
    - Callgrind outputs (summarized) and `TMPDIR`.
- **Kept:** the worktree and branch.

The packet keeps sources, scripts, compressed raw reports and summaries. It
holds no binaries, `perf.data`, Callgrind outputs or corpora.

[Evidence packet and replay instructions](results/change-0769/README.md).
