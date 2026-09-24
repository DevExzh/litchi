# 0767 — CFB open and read clear their chain maps in proportion to the chains: same verdicts, flat work per stream, opens of 10,000 streams 3–32% faster

Status: retained, implemented in `litchi-cfb`. `performance_claim: none`.
Paired timings, instruction and cycle counts, Callgrind counts and
allocation counts are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF work stays deferred until that
goal completes; iWork is excluded.

Base `1d1044e3ac`; branch `perf/0767-cfb-reparse-linear`.

- `af72192785`: the reusable chain maps keep every bit clear between walks;
  `OleFile` keeps one chain scratch for its reads; tests.
- `573554acfe`: the restoration zeroes whole words and only for chains under
  an eighth of the table's words, which removed most of the first version's
  instruction cost on small tables. The tests draw tables that take both
  restorations. This is the measured candidate.
- `866229144e`: a doc comment corrected (the retained bound for mini
  streams). The rebuilt probe's `.text`, `.rodata` and `.data.rel.ro` are
  byte-identical to the measured binary's.

## Result

Opening a compound file and reading its streams cleared, per stream, a
visited map sized to the whole FAT or MiniFAT. Record
[0749](0749-cfb-reuse-plan-validation.md) found the first of these terms in
the structural reparse; this record found a second one in `open_stream`.
Both now cost time proportional to the chain walked. Every verdict, check
and error string is unchanged, and the work per stream no longer grows with
the stream count.

| Lane (12 paired layouts, core 28) | Base p50 | Cand p50 | Paired Δ p50 [95% CI] |
|---|---:|---:|---:|
| `OleFile::open`, v3, 10,000 × 2,000-byte (mini) streams | 4,539 µs | 3,137 µs | **−31.26%** [−31.76, −30.49] |
| `OleFile::open`, v4, 10,000 mini streams | 4,514 µs | 3,054 µs | **−32.47%** [−33.04, −31.83] |
| `OleFile::open`, v3, 10,000 × 5,000-byte (regular) streams | 2,348 µs | 1,888 µs | **−20.20%** [−20.80, −19.26] |
| `OleFile::open`, v4, 10,000 regular streams | 1,787 µs | 1,722 µs | −3.41% [−5.02, −2.35] |
| Read every stream (`open_stream`), v3, 10,000 mini streams | 7,720 µs | 4,604 µs | **−39.22%** [−40.24, −36.89] |
| Read every stream, v3, 10,000 regular streams | 6,064 µs | 5,295 µs | −12.94% [−15.06, −11.63] |
| `SharedOleFile::open_owned`, v3, 10,000 mini streams | 4,003 µs | 2,601 µs | **−34.86%** [−35.28, −34.10] |
| Common-editor open and capture (`Editor::open`), v3, 3,000 mini streams | 3,524 µs | 2,834 µs | **−19.93%** [−21.16, −18.89] |
| Reuse `write_to`, 45543.ppt (0749's lane) | 22.4 µs | 22.5 µs | +0.48% [+0.00, +1.85] |
| `OleFile::open`, 45543.ppt / FloatingPictures.doc / WithCheckBoxes.xls | 3.3 / 5.3 / 4.5 µs | 3.3 / 5.4 / 4.5 µs | +0.92% / +1.42% / −0.44% |

- **The work per stream is flat.** Native instructions per stream of an
  open are 7,296, 7,285 and 7,285 at 1,000, 3,000 and 10,000 v3 mini
  streams, against the base's 7,181, 7,387 and 8,157. Callgrind puts the
  stream-allocation validation (A5) at ×3.00 and ×3.33 for three and 3.33
  times the streams, against the base's ×6.57 and ×9.67.
- **Reads allocate one buffer fewer per stream.** At 10,000 mini streams a
  read of every stream allocates 42.4 MB in 50,038 calls instead of
  444.7 MB in 100,037. Opens allocate exactly what they did.
- **Real files do essentially the same work.** An open of the three
  fixtures above runs +74 to +158 instructions (+0.10% to +0.16%). Their timings stay
  within +1.4%, and no small-file case has a flag above +5% in its median.
- **Output and verdicts are identical.** Every probe process of both arms
  published the same output for its case, including the sealed 0728 digest
  `545ed7e5…` for the 45543.ppt removal. 30,400 byte-level faults over 104
  files give byte-identical verdict streams from the two builds.

## The two super-linear terms

Callgrind of one timed owner on the base, generated v3 files
(`profiles/` in the packet):

| Owner | Streams | Owner Ir | Of which the term | Term's share |
|---|---:|---:|---:|---:|
| `OleFile::open`, mini streams | 1,000 | 11.8 M | 5.2 M (`SectorChainScratch::collect_exact`) | 43.9% |
| | 3,000 | 59.2 M | 39.5 M | 66.7% |
| | 10,000 | 474.6 M | 411.6 M | 86.7% |
| Read every stream, mini streams | 1,000 | 12.6 M | 7.0 M (`collect_sector_chain`) | 55.6% |
| | 10,000 | 490.4 M | 431.0 M | 87.9% |

Callgrind counts its own byte loops for `memset`, which inflates both terms;
the native counters below measure the same terms more mildly, and the
timings show what they cost.

- **A5, on every open.** `validate_stream_allocations` walks every stream's
  chain through a reusable `SectorChainScratch`. Its `prepare_visited`
  cleared the scratch's whole visited map, one bit per FAT or MiniFAT
  entry, before each stream: streams × table words. 0749 measured this term
  (×6.93 from 1,000 to 3,000 streams) and left it for its own record,
  because it runs on every `OleFile::open`, including `SharedOleFile`'s
  open, the Reuse-plan validation's reparse, `SourceLayout::parse` and the
  sequential writer's validation.
- **`open_stream`, on every read.** `read_stream_from_fat` and
  `read_stream_from_minifat` collected the chain with `collect_sector_chain`,
  which allocated and zeroed a fresh table-sized map for every call. Reading
  every stream of a file, as the common editor's capture does, was therefore
  quadratic too.

## What was changed

`crates/litchi-cfb/src/file.rs`:

- `CheckedBitSet` gains three helpers shared by both reusable chain
  collectors:
  - `cover` makes the map cover a table by growth only: new words arrive
    clear, and existing words are clear by the invariant.
  - `clear_walk` restores the all-clear invariant after a successful walk,
    whose only set bits are the sectors it recorded. It zeroes the whole word
    holding each recorded sector when the chain is shorter than an eighth of
    the table's words (`CLEAR_BY_SECTOR_RATIO`), and clears the table's
    words otherwise. So each walk writes at most eight words per chain
    sector, and never more than the table.
  - `clear_table` clears the table's words, and it is how a failed walk
    restores the map.
  - The per-bit `remove` 0749 added is removed; nothing else used it.
- `SectorChainScratch` (A5): `prepare_visited` now only covers the table,
  and `collect_exact` restores the map after the walk. It uses
  `clear_walk` on success, and `clear_table` then `reset` on error. The
  walk itself, its checks, their order and every error string are
  untouched.
- `EndChainScratch` (0749's) uses the same helpers. It behaves as before,
  except that the restoration zeroes words and prefers the bulk clear up to
  the ratio.
- `OleFile` gains a private `stream_chain: EndChainScratch`, which
  `open_stream`'s chain collection reuses:
  - `read_stream_from_fat` moves it out while the batched read borrows the
    reader, runs the old body as `read_fat_chain`, and puts it back on every
    path;
  - `read_stream_from_minifat` collects into it directly.
  - It is empty until the first read. It then retains one bit per entry of
    the larger table read and one `u32` per sector of the longest chain
    read. That is 1/128 (512-byte sectors) or 1/1024 (4096-byte sectors) of
    a FAT-chained stream's bytes, and at most 64 entries for a validated
    mini stream's chain.
- `collect_sector_chain`, the allocating form, is now `#[cfg(test)]`: the
  oracle the tests hold the scratch to.
- `visited_map_work`: a test-only per-thread count of the words the chain
  maps write. Outside tests it compiles to nothing.

New tests: `crates/litchi-cfb/src/chain_work_tests.rs`, a child module of
`file`; 13 test-only `OleFile` literals gain the new field.

There is no public API change, no new dependency, no `unsafe`, no new or
changed limit, and no change to planning, emission or any published byte.
The save functions record 0761 is changing (`writer/core.rs`,
`writer/sequential.rs`, `overlay.rs`) are untouched.

## Why the verdicts cannot differ

- **A reused map reads as a fresh one.** Between walks every bit is clear,
  and `cover` sets `bit_len` to the table's length. `contains` and `insert`
  consult only the table's bits, so a walk sees exactly the map a fresh
  table-sized allocation gave it, whichever table the scratch saw before.
- **The walks are the same code.** `collect_exact`'s loop and
  `walk_chain_to_end` are unchanged, so every index, cycle, marker, exact
  length and bounds check runs in the same order with the same error
  string. `validate_stream_allocations`' ownership checks (the
  `claimed_mini_sectors` map and `claim_chain`) are untouched.
- **The restoration is exact.**
  - In `collect_exact`, each `insert` is immediately followed by the
    infallible `push` of its sector into the reserved chain. In
    `walk_chain_to_end`, a successful walk has recorded every sector it
    marked.
  - So after a success, every set bit belongs to a recorded sector, and
    zeroing a recorded sector's whole word clears only recorded bits.
  - After any failure, including an allocation failure between
    `walk_chain_to_end`'s `insert` and its `try_push`, the table's words are
    cleared.
- **The reader's retained scratch cannot leak state.** It is restored before
  `collect` returns, on success and on error, and `read_stream_from_fat`
  puts it back whatever the read's outcome. A panic would leave a fresh
  default scratch in its place, which is also all-clear.
- **Allocation errors keep their types and labels.** Growth reserves
  fallibly under the same `"sector-chain map"` and `"sector-chain entries"`
  labels. The one difference: a buffer the scratch no longer allocates per
  call can no longer fail to allocate.

## Other terms checked

The whole open and read path was profiled at 1,000, 3,000 and 10,000
streams in both geometries (`scaling/` in the packet).

- **The rest of the open is linear.** In the candidate every open owner
  scales ×2.96 to ×3.00 for three times the streams and ×3.12 to ×3.23 for
  3.33 times. This covers the header and DIFAT, the FAT, the directory
  load, `validated_directory_entries`' tree walk, the storage tree build,
  the MiniFAT, the range-lock scan, A5 and the partition check. The
  directory, MiniFAT and root-chain collections each allocate one
  table-sized map per open, not per stream.
- **Reads keep one logarithmic term: the path lookup.** `find_entry`
  descends the MS-CFB sibling tree by name comparison, so its cost per
  stream is 1,441, 1,600 and 1,799 Ir at 1,000, 3,000 and 10,000 streams,
  about one tree level per doubling.
  - It is why a read of every stream grows 5% in instructions per stream
    from 1,000 to 10,000 mini streams (6,971 → 7,349). Without it, the reads
    scale ×3.02 and ×3.33.
  - A list-shaped sibling tree, which the reader accepts because deployed
    writers produce unbalanced trees, would make each lookup linear in its
    siblings.
  - A per-open name index would remove that, at a cost every open would pay.
    Owner trade-off 3 of [0652](0652-owner-decisions-for-the-third-wave.md)
    says not to make every input pay for a minority, so it is not done here.
- **The time per stream still grows at 10,000 streams in both arms.** That
  is memory, not work: the reads copy 20 to 50 MB of stream bytes and a 20 MB
  mini stream, larger than the caches. The instructions per stream stay flat.
- **`SharedOleFile` reads, the 0749 `StreamComparer` and the Reuse planner**
  were already linear or O(n log n):
  - the shared reader walks exact chains without a visited map;
  - the comparer uses `EndChainScratch`;
  - the planner's `high_water` scan runs a fixed number of rounds, and its
    maps are ordered B-trees.
  - `SourceLayout::parse` is one open plus one bounded walk per regular
    stream.
- **A pre-existing gap, identical in both builds (not changed; see
  follow-ups).** The verdict lane found that A5 admits a mini stream in the
  last mini sector of a root mini stream whose size is not a multiple of 64.
  `open_stream` then refuses that stream with "Mini sector out of bounds".
  This accounts for all 52 read errors after a successful open in the lane.

## Authority and constraints

- **ADR 0006:** validation never mutates, refusals stay typed with the same
  strings, and nothing is repaired. Every check of 0749's A1–A6 runs as
  before.
- **ADR 0005:** no ambient behaviour, threads or I/O. The retained buffers
  are bounded by the tables and chains the reader already holds, as
  described under what was changed.
- **ADR 0003:** no published byte, patch or inverse changes.
- **ADR 0026:** directory validation stays in `litchi-cfb`, unchanged.
- **GOAL, "LEGACY CFB-SPECIFIC WORK":** every validation, ownership, cycle,
  overlap, FAT, MiniFAT, directory and truncation check is preserved.
- **Change 0652's standing trade-offs:**
  - Correctness first: the verdicts are held to fresh-map oracles in-crate
    and to the base build byte for byte.
  - The common benign path: the first version's per-bit restoration cost
    small tables 0.2–0.3% more instructions per open. The refinement keeps
    small tables on the bulk clear, which leaves +0.10% to +0.16%.
- **Change 0758** decides nothing this record uses. No ODF or iWork crate is
  touched.

## Proof obligations and tests

`chain_work_tests.rs`:

- **`sector_chain_scratch_matches_fresh_maps_over_random_call_sequences`**
  and **`end_chain_scratch_matches_fresh_maps_over_random_call_sequences`.**
  Each drives one scratch through 4,800 collections: 400 tables, long and
  short, each shuffled into chains with up to three corrupted links (joins,
  loops, early ends, `FREESECT`, markers, indices past the table). Each
  collection must equal the allocating helper's result or error string, and
  the map must be all-clear after every call. Both restorations are taken
  at least 150 times.
- **`stream_allocation_validation_matches_the_fresh_map_form_on_faulted_files`.**
  12 generated v3/v4 files mix empty, mini and regular streams in three
  sibling trees, with 150 seeded fault sets each (1,800 cases). The faults
  are FAT and MiniFAT links, shared or moved starts, whole shared
  allocations, sizes off by a byte, a mini sector or a sector, the cutoff
  crossed, and root moves and resizes. `validate_stream_allocations` must
  equal a copy of its fresh-map form in result, error string, claimed
  sector roles and recorded root chain. The test requires the cases to
  reach these refusals, and they do: the three cycle kinds, both ownership
  overlaps, ends early and late, index, marker and start-marker errors,
  mini sectors outside the root mini stream, and non-terminated empty
  chains.
- **`reads_through_one_reader_match_a_fresh_reader_per_stream_on_faulted_tables`.**
  The faults are applied after a validating open, so the reads meet chains
  the open never saw. Reading every stream through one reader must equal a
  fresh reader per stream, stream for stream.
- **The work tests.** `validating_open_clears_chain_maps_in_proportion_to_the_chains`
  and `reading_every_stream_collects_chains_in_proportion_to_the_chains`
  count the words written at 64, 256 and 1,024 streams, mini and regular,
  v3 and v4. The count must stay within the tables' words plus eight per
  chain sector, within a quarter of the per-stream table clearing at 1,024
  streams (open), and flat per stream from 64 to 1,024.
- **`a_reader_reuses_its_chain_buffers_across_reads_and_errors`.** A looped
  regular chain fails with the reader's own error, and the reader then
  reads every stream correctly in the same buffers.

**Mutation checks.**

- A restoration that never zeroes fails both random-sequence tests, 0749's
  `scratch_to_end_clears_only_what_it_set_and_reuses_its_buffers`, and four
  existing overlay, stream-compare and writer tests.
- The base's per-call full clear passes every verdict test and fails
  exactly the two work tests, 557,716 words against a bound of 295,498.

**The cross-build verdict lane** (probe mode `verdicts`, packet
`verdicts/`):

- 300 seeded byte-level fault sets per file on the 98 `test-data/ole`
  fixtures and 45543.ppt and 41246-1.ppt; 100 per file on four generated
  1,000-stream files.
- The faults are FAT and MiniFAT links, and stream and root directory start
  and size fields.
- Each case records `OleFile::open`'s verdict, then every stream read on
  that reader, and `SharedOleFile::open_owned` with its reads.
- The 104 outputs of the base and candidate builds are byte-identical:
  30,400 faulted cases (2,588 accepted, 27,812 refused across 51 distinct
  refusal messages, numbers masked) and 53,257 stream reads.

## Method

- **Arms.** Base `1d1044e3ac` (a detached checkout) and candidate
  `573554acfe`.
  - Both were built with identical cargo commands, flags and features under
    rustc 1.95.0 (`binaries.json` has the commands and every SHA-256).
  - The probe manifests differ only in the source tree their dependencies
    point at, and `--remap-path-prefix` maps both trees to `/litchi`. A
    repeated build compiled nothing in either arm.
  - The gitignored root `Cargo.lock` copied into the worktrees predates the
    spec-gap merge. `cargo metadata --offline` refreshed it once, adding
    edges without changing a version, and both legs used that one file.

  | Binary | Base | Candidate |
  |---|---|---|
  | probe | `ca80d7ef…79f8` | `733049d1…0ad4` |
  | probe_alloc | `2a5c6033…0c27` | `57f77482…1a5e` |
  | litchi-perf-baseline | `11f77d34…c79f` | `784486e1…6ad6` |

- **The probe** (`probe/`, public APIs only). Modes:
  - `cfb-open`, `shared-open` and `editor-open`: the three opens;
  - `cfb-read-all`: every stream on a reader opened untimed for the owner;
  - `cfb-open-read`: 0749's reader control;
  - `cfb-write-reuse` and `cfb-write-rewrite`: 0749's CFB-only writes, the
    Rewrite one a control whose timed region calls no reader code;
  - `generate`: 0749's generator. It writes v3 and v4 files of 1,000, 3,000
    and 10,000 streams of 2,000 bytes (mini) or 5,000 bytes (regular);
    digests are in `inputs.json`;
  - `verdicts`: the fault lane.
- **Harness selectors** `cfb_open` and `cfb_list_streams` (shapes tiny,
  many-small, few-large, wide-root; incompressible payload) and
  `doc_semantic_open`, `xls_semantic_open`, `ppt_semantic_open` (the tiny
  and large OLE2 semantic corpora; the harness has no medium one).
- **Timing.** 12 rounds of 39 cases, ABBA across rounds with rotated case
  order. Each binary starts through a symlink 8 bytes longer each round
  (0745's layout sampling). Every process is pinned with `taskset -c 28`,
  and other agents built on other cores. 936 processes ran and none failed.
  - Samples per process: at least 20 after at least 3 warmups for the
    10,000-stream owners, and up to 5,000 for microsecond owners.
  - Statistics are 0749's: process p50 as the midpoint median, a comparison
    as the median of per-round paired changes, and a 10,000-resample
    bootstrap interval (seed 767).
- **Counters.** Per-owner `instructions:u` and `cycles:u` come from two
  processes of different sample counts at three argv[0] layouts; their
  difference removes setup. Read-all owners include the untimed per-owner
  open, so the per-stream read figures subtract the open owner's.
- **Callgrind.** One owner, collection toggled on `timed_owner`, per arm,
  mode and generated file (48 runs).
- **Allocation.** One counting-allocator process per case and arm, five
  owners each. Every owner of a process was identical.

## Results

Selected rows; `summary.md` has every row, flag, counter, Callgrind row and
allocation row.

**Scaling, per stream** (native `instructions:u` per owner ÷ streams, and
p50 time ÷ streams):

| Input | Streams | Open instr/stream, base → cand | Open ns/stream | Read-all instr/stream | Read-all ns/stream |
|---|---:|---:|---:|---:|---:|
| v3 mini | 1,000 | 7,181 → 7,296 | 370 → 359 | 8,841 → 6,971 | 420 → 343 |
| | 3,000 | 7,387 → 7,285 | 346 → 313 | 9,175 → 7,170 | 456 → 350 |
| | 10,000 | 8,157 → 7,285 | 454 → 314 | 10,109 → 7,349 | 772 → 460 |
| v3 regular | 1,000 | 3,846 → 3,877 | 237 → 234 | 5,959 → 4,543 | 321 → 266 |
| | 3,000 | 3,909 → 3,870 | 200 → 186 | 6,232 → 4,724 | 384 → 310 |
| | 10,000 | 4,155 → 3,871 | 235 → 189 | 6,643 → 4,901 | 606 → 529 |
| v4 mini | 1,000 | 6,984 → 7,093 | 338 → 330 | 8,565 → 6,782 | 412 → 334 |
| | 3,000 | 7,162 → 7,053 | 339 → 302 | 9,014 → 6,987 | 462 → 346 |
| | 10,000 | 7,927 → 7,055 | 451 → 305 | 9,971 → 7,162 | 762 → 530 |
| v4 regular | 1,000 | 3,329 → 3,341 | 195 → 195 | 4,364 → 4,121 | 272 → 259 |
| | 3,000 | 3,340 → 3,329 | 171 → 168 | 4,549 → 4,288 | 370 → 363 |
| | 10,000 | 3,397 → 3,337 | 179 → 172 | 5,123 → 4,475 | 574 → 527 |

At 1,000 mini streams the candidate's open runs about 115 more
instructions per stream but fewer cycles (−1.5% for v3). The base cleared
each stream's 500-word map with one bulk `memset`, which runs few
instructions and many cycles. The candidate zeroes 32 words one by one.

**Timing** (Δ is the median paired change of p50 [95% CI]):

| Case | Base p50 µs | Cand p50 µs | Δ p50 [95% CI] | Δ mean | Δ p95 |
|---|---:|---:|---:|---:|---:|
| `open-v3-mini-1000` / `-3000` / `-10000` | 370.2 / 1,038.8 / 4,539.1 | 359.3 / 938.3 / 3,137.0 | −3.36% / −10.56% / −31.26% | −2.94% / −10.59% / −31.30% | −3.37% / −10.57% / −31.30% |
| `open-v3-regular-1000` / `-3000` / `-10000` | 237.4 / 600.2 / 2,348.0 | 234.1 / 558.9 / 1,887.7 | −1.61% / −7.19% / −20.20% | −1.91% / −7.05% / −20.29% | −2.31% / −7.14% / −20.57% |
| `open-v4-mini-1000` / `-3000` / `-10000` | 338.5 / 1,018.4 / 4,513.9 | 330.3 / 906.2 / 3,054.1 | −2.30% / −10.71% / −32.47% | −2.40% / −10.75% / −32.50% | −2.51% / −10.95% / −32.61% |
| `open-v4-regular-1000` / `-3000` / `-10000` | 194.9 / 512.1 / 1,787.1 | 195.5 / 505.2 / 1,721.5 | +0.85% / −1.24% / −3.41% | +0.33% / −1.25% / −3.46% | +0.48% / −1.08% / −4.89% |
| `readall-v3-mini-1000` / `-3000` / `-10000` | 420.2 / 1,367.0 / 7,720.0 | 343.4 / 1,049.5 / 4,604.1 | −18.89% / −24.33% / −39.22% | −18.96% / −24.18% / −38.97% | −18.83% / −23.93% / −37.32% |
| `readall-v3-regular-1000` / `-3000` / `-10000` | 320.6 / 1,151.5 / 6,064.1 | 266.3 / 930.9 / 5,294.9 | −16.64% / −16.74% / −12.94% | −16.31% / −17.01% / −13.03% | −16.93% / −17.88% / −12.41% |
| `readall-v4-mini-1000` / `-3000` / `-10000` | 412.5 / 1,384.8 / 7,619.0 | 334.4 / 1,037.7 / 5,297.4 | −18.73% / −24.98% / −31.20% | −18.65% / −24.89% / −31.72% | −18.64% / −26.45% / −31.18% |
| `readall-v4-regular-1000` / `-3000` / `-10000` | 271.6 / 1,111.2 / 5,745.0 | 259.2 / 1,088.4 / 5,274.1 | −4.43% / −2.81% / −6.74% | −4.16% / −4.17% / −6.67% | −5.15% / −6.20% / −6.82% |
| `shared-open-v3-mini-10000` / `-regular-10000` | 4,002.6 / 2,329.0 | 2,601.0 / 1,864.6 | −34.86% / −19.97% | −34.91% / −20.00% | −34.79% / −19.59% |
| `editor-open-v3-mini-1000` / `-3000` / `-10000` | 965.2 / 3,523.8 / 18,182.9 | 801.7 / 2,834.0 / 17,084.5 | −16.83% / −19.93% / −4.54% [−16.07, −2.37] | −17.04% / −20.09% / −4.95% | −17.00% / −19.90% / +5.78% |
| `write-reuse-45543` | 22.4 | 22.5 | +0.48% [+0.00, +1.85] | +0.44% | +0.54% |
| `write-reuse-v3-mini-3000-grow` | 6,372.6 | 6,271.8 | −3.78% [−5.35, −2.72] | −4.13% | −3.52% |
| `open-45543` / `-floating` / `-checkboxes` | 3.3 / 5.3 / 4.5 | 3.3 / 5.4 / 4.5 | +0.92% / +1.42% / −0.44% | +0.85% / +1.49% / +0.86% | +0.44% / +1.49% / −0.33% |
| `openread-45543` / `editor-open-45543` | 10.6 / 98.1 | 10.4 / 97.3 | −2.86% / −0.81% | −2.87% / −0.70% | −2.80% / −0.76% |
| `ctl-write-rewrite-45543` (control) | 8.7 | 8.7 | +0.24% [−4.11, +5.80] | +0.24% | −0.73% |
| `cfb_open` tiny / many-small / few-large / wide-root | 1.5 / 77.5 / 68.1 / 483.7 | 1.5 / 78.4 / 68.9 / 493.3 | +1.67% / +0.65% / +1.12% / +1.83% | | |
| `cfb_list_streams` tiny / many-small / few-large / wide-root (controls) | 0.13 / 9.5 / 0.16 / 73.5 | 0.15 / 9.7 / 0.16 / 73.4 | **+15.38%** / +1.70% / −1.56% / −0.03% | | |
| `doc_semantic_open` tiny / large | 6.1 / 311.3 | 6.0 / 283.2 | −1.83% / −8.53% | | |
| `xls_semantic_open` tiny / large | 10.3 / 1,330.0 | 10.3 / 1,309.1 | +0.05% / −1.27% | | |
| `ppt_semantic_open` tiny / large | 5.9 / 12.9 | 5.9 / 12.9 | −0.59% / −0.44% | | |

**Per-owner counters** (median of three layouts; `summary.md` has every
case):

| Case | instructions:u base → cand | Δ | cycles:u base → cand | Δ |
|---|---:|---:|---:|---:|
| `open-v3-mini-10000` | 81.57 M → 72.85 M | −10.69% | 20.17 M → 13.98 M | −30.68% |
| `open-v3-regular-10000` | 41.55 M → 38.71 M | −6.83% | 10.67 M → 8.53 M | −20.07% |
| `readall-v3-mini-10000` (with its untimed open) | 182.67 M → 146.35 M | −19.88% | 52.25 M → 33.13 M | −36.59% |
| `editor-open-v3-mini-10000` | 207.93 M → 166.82 M | −19.77% | 79.82 M → 56.83 M | −28.80% |
| `open-v3-mini-1000` | 7,181,129 → 7,296,165 | +1.60% | 1,446,940 → 1,425,525 | −1.48% |
| `write-reuse-45543` | 383,867 → 379,868 | −1.04% | 103,931 → 102,355 | −1.52% |
| `open-45543` | 74,726 → 74,800 | +0.10% | 14,875 → 15,003 | +0.86% |
| `open-floating` | 113,982 → 114,140 | +0.14% | 24,061 → 24,334 | +1.14% |
| `open-checkboxes` | 89,376 → 89,520 | +0.16% | 20,325 → 20,584 | +1.27% |
| `cfb_open` tiny / many-small / few-large / wide-root | | +0.15% / +0.16% / +0.00% / +0.04% | | |
| `cfb_list_streams` (all shapes) | | −0.31% to +0.03% | | |
| `doc/xls/ppt_semantic_open` (both shapes) | | −0.11% to +0.01% | | |

The harness counters include each iteration's untimed setup. For the
semantic opens that setup dominates: about 29 M, 131 M and 6 M
instructions per owner. So they show that instructions are equal, and
their cycles are not a measure of the timed region.

**Allocation** per owner: every open allocates exactly what it did. The
reads drop one allocation per stream and the chain vector's regrowth:

| Case | Bytes base → cand | Calls | Peak live bytes |
|---|---:|---:|---:|
| `readall-v3-mini-10000` | 444,717,368 → 42,357,368 | 100,037 → 50,038 | 21,435,376 → 21,477,376 |
| `readall-v3-regular-10000` | 181,707,416 → 51,320,456 | 90,020 → 50,024 | 706,224 → 711,224 |
| `editor-open-v3-mini-10000` | 501,892,002 → 99,532,002 | 140,100 → 90,101 | 71,676,384 → 71,942,400 |
| `openread-45543` | 406,729 → 404,745 | 95 → 76 | 322,511 → 322,607 |
| every `open-*`, `shared-open-*`, write and control case | unchanged | unchanged | unchanged |

The peak rises by the retained chain buffers, from 96 bytes (45543.ppt) to
266 KB (the 10,000-stream capture, 0.37% of its peak). Bytes retained by a
reader after its reads rise likewise, by 302 KB beside a 20 MB cached mini
stream at 10,000 mini streams, and by 13 KB at 10,000 regular streams.

## The first version and why it was refined

The first version (`af72192785`) restored a map by clearing each recorded
bit, one read-modify-write per sector, whenever the chain was shorter than
the table's words. Its own matrix (8 rounds, packet `matrix-v1/` and
`counters-v1/`) measured the same large gains. It also ran more
instructions where the old code cleared a short table with one small
`memset`:

| Case | First version, instructions Δ | Retained, instructions Δ |
|---|---:|---:|
| `open-45543` / `open-floating` / `open-checkboxes` | +0.30% / +0.21% / +0.22% | +0.10% / +0.14% / +0.16% |
| `open-v3-mini-1000` / `open-v4-mini-1000` | +2.92% / +2.92% | +1.60% / +1.55% |
| harness `cfb_open/many-small` | +2.04% | +0.16% |

Callgrind of the 45543.ppt open located it in `collect_exact`'s restore
loop. The retained version keeps small tables on the bulk clear and zeroes
whole words (exact, because only recorded bits are set), so each sector
costs one store. The remaining +74 instructions on 45543.ppt, a file of five
streams, are the restoration's bookkeeping, about 15 per stream; the
`memset` count is the base's (10,564 Ir in both). The small files' cycles stayed about +1% in both
versions while their instructions halved. This record therefore reads that
cycle difference as code placement, not as the added instructions.

## Regression flags (every change above +5%)

The retained matrix has 83 per-round flags (a round's paired p50, mean or
p95 above +5%); `summary.md` lists each with its values. Round 2 holds 12 of
them across unrelated cases, which points at a host disturbance.

- **`cfb_list_streams/tiny`, a harness control: +15.38% [+7.69, +15.39]
  median p50, with 26 flags in rounds 3 to 11.** This is the one case whose
  median is above +5%.
  - It times `list_streams` of a three-stream file, 0.13 µs, which this
    change does not touch. The difference is 20 ns (130 → 150 ns).
  - Its instructions per owner are equal (4,984 → 4,982).
  - It is code placement in the harness binary, like 0749's
    `cfb_open/few-large` flag, and is kept as a flagged limitation.
- **`ctl-write-rewrite-45543`, a control, 8 flags** (p50 +5.8% to +6.7% in
  rounds 2, 5, 7 and 10; median +0.24% [−4.11, +5.80]).
  - Its timed region calls no reader code.
  - Callgrind attributes its −2.33% instructions to `realloc` of the
    output buffer, which extends in place or moves depending on the heap
    state the untimed setup leaves: 870 K → 717 K Ir in `realloc`, the
    whole difference.
  - The setup reads streams through the changed reader, which now
    allocates less.
- **`doc_semantic_open/large`, 7 flags** (p95 +22% to +32% in rounds 8 to
  11). Its median p50 is −8.53%, with equal instructions and 592 → 267 page
  faults per owner. That matches the allocator-state behaviour 0749 tested;
  it was not re-tested here.
- **`editor-open-v3-mini-10000`, 6 p95 flags** (+7.7% to +24.5%). Its median
  p50 is −4.54% with a wide interval [−16.07, −2.37]. The 18 ms owner
  allocates 502 MB (base) or 99.5 MB (candidate) per capture.
- **`open-checkboxes`, 5 flags** in rounds 2, 3 and 5, including round 2's
  mean +21.6%; its median is −0.44%.
- **`write-reuse-45543`, 4 flags:** rounds 1–3 mean or p95, up to +23.6%
  p95; median +0.48%, instructions −1.04%.
- **Three flags each:** `cfb_list_streams/few-large` (rounds 3, 6 and 7:
  10 ns steps on a 150–190 ns call), `cfb_list_streams/many-small` (round
  8), `open-floating`
  (rounds 6 and 7), `open-v3-regular-1000` (round 2) and
  `readall-v4-regular-3000` (rounds 1 and 9).
- **Two flags each:** `cfb_open/few-large` (round 6), `editor-open-45543`
  (rounds 1 and 8) and `readall-v3-regular-1000` (round 2).
- **Single flags:** `cfb_open/many-small`, `cfb_open/tiny`,
  `open-v3-regular-3000`, `open-v4-mini-1000`, `open-v4-regular-1000` and
  `write-reuse-v3-mini-3000-grow`.

## What is not claimed

- No registered speedup: `performance_claim: none`.
- **No gain for ordinary files.** Real fixtures run +0.10% to +0.16%
  instructions per open, and their timings stay within +1.4%. The gains are
  for files with thousands of streams, such as the generated inputs here,
  which are not a corpus of real documents.
- **No claim that reading every stream is linear.** The instructions per
  stream grow with the sibling-tree depth, and the time per stream grows at
  10,000 streams with the working set in both arms.
- **No claim about the harness selectors' cycles.** Their per-owner cycle
  deltas are dominated by untimed per-iteration setup.
- **No claim for `doc_semantic_open/large`'s −8.53%,** which follows page
  faults, not work.
- **No change to the Reuse write's planning or emission,** and no claim for
  the Reuse write on 45543.ppt (+0.48%, instructions −1.04%).
- **Scope:** warm, in-memory, serial, on one host; no cold-I/O, concurrency,
  RSS or cross-platform result. Peak live bytes are allocator ownership, not
  RSS.

## Follow-up candidates found, not changed

- **The partial last mini sector.** A5 admits a mini stream in the last
  mini sector of a root mini stream whose size is not a multiple of 64: its
  capacity rounds up. `open_stream` and the shared reader's whole-stream
  read then refuse it with "Mini sector out of bounds", while range reads
  check only the bytes they need. The refusal is typed and the same in both
  builds. Aligning the three would change a verdict, so it needs its own
  record.
- **The path lookup's logarithmic term,** and a list-shaped sibling tree's
  linear one, remain in every read by path. A name index per storage is the
  candidate if many-stream reads matter; owner trade-off 3 weighs against
  building it on every open.
- **The per-owner cycles of harness selectors** are not usable as evidence
  where untimed setup dominates the iteration.

## Verification

`gates.txt` in the packet lists every command and exit code. They ran one
after another after all measurement, on `573554acfe`, with
`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=6` and `TMPDIR` under the scratch
directory. On `866229144e`, the comment-only follow-up, fmt, Clippy (all
targets), rustdoc and the `litchi-cfb` tests (501 passed) were rerun and
pass.

- `cargo fmt --all --check`: pass.
- `cargo check --all-targets --locked --offline`: pass, for three sets:
  - `litchi-cfb` and its eight direct dependents (`litchi-ole-common`,
    `litchi-doc`, `litchi-ppt`, `litchi-xls`, `litchi-vba`, `litchi-ograph`,
    `litchi-crypto`, `litchi-sign`);
  - the OOXML crates that reach it through `litchi-crypto` (`litchi-opc`,
    `litchi-ooxml-common`, `litchi-drawingml`,
    `litchi-spreadsheet-drawing`, `litchi-docx`, `litchi-pptx`,
    `litchi-xlsx`, `litchi-xlsb`);
  - the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt`.
- Clippy `-D warnings` on the `litchi-cfb` library and on all its targets:
  pass.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps`: pass. It
  also passes with `--document-private-items`.
- Tests:

  | Crate | Passed | Failed | Ignored |
  |---|---:|---:|---:|
  | `litchi-cfb` | 501 | 0 | 1 |
  | `litchi-ole-common` | 255 | 0 | 1 |
  | `litchi-doc` | 1,293 | 0 | 13 |
  | `litchi-ppt` | 1,249 | 0 | 11 |
  | `litchi-xls` | 1,611 | 0 | 1 |
  | `litchi-vba` | 47 | 0 | 0 |
  | `litchi-ograph` | 70 | 0 | 0 |
  | `litchi-crypto` | 43 | 0 | 0 |
  | `litchi-sign` | 25 | 0 | 0 |
  | Facade (`litchi`, all OLE2/OOXML features) | 382 | 0 | 7 |

  The ignored tests are pre-existing.
- Crate boundaries: pass (65 packages, 244 declarations, 11 existing debt
  items).
- Structural perf-claims check: pass (10 claims).
- `non_iwork_gate verify`: fails with "workspace package inventory mismatch
  (unexpected: litchi-xldm)". This is the known failure at base
  `1d1044e3ac` that the wave briefing lists, and it is being fixed on
  another branch.

The harness is unchanged, so its tests and the coverage validator were not
run. The briefing's other known failures (two harness tests, the facade's
`clippy::unit_arg` lints, iWork crates under workspace-wide
`--all-targets`) are outside these commands. No shared log, claim registry
or coverage index is edited; `log-sections.md` in the packet has the
paste-ready sections.

## Cleanup

`cleanup.json` records every removal.

- **Removed after committing this record and packet:**
  - `targets/0767` and `targets/0767-before`: the debug gate build, the
    probe builds and the harness builds;
  - the `0767-before-src` checkout, with `git worktree remove --force`;
  - `scratch/0767`: staged binaries (hashes in `binaries.json`), generated
    inputs (regenerable, digests in `inputs.json`), argv0 symlinks, raw lane
    outputs (bundled in the packet), Callgrind outputs (summarized) and
    `TMPDIR`.
- **Kept:** the worktree and branch.

The packet keeps sources, scripts, compressed raw reports and summaries
(4.3 MB). It holds no binaries, `perf.data`, Callgrind outputs or corpora.

[Evidence packet and replay instructions](results/change-0767/README.md).
