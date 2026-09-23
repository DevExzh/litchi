# 0749 — Reuse-plan validation compares streams in place: same verdicts, a linear readback, the CFB Reuse write 4–41% faster

Status: retained, implemented in `litchi-cfb`. The first version (one
production commit and one test commit) was quadratic in mini streams; the
review round fixed that in three further commits and re-measured.
`performance_claim: none`. Paired timings, instruction and cycle counts and
allocation counts are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF work stays deferred until that
goal completes; iWork is excluded.

Base `ab29ac6291` (the branch tip with record 0745); branch
`perf/0749-cfb-reuse-plan-validation`.

- `897d24fd0f`: `ReusePlan::validate` reads streams back in place.
- `776ad05175`: the fault-injection tests also cover version 4 (4096-byte
  sector) plans.
- `56548fd06e` (review round): the readback loads the root mini stream once
  per validation, compares contiguous mini sectors as one range and clears
  its chain map in proportion to the chains, so it is linear in the streams;
  the reader and the comparison share one chain walker and one run
  segmentation; `examine_at` fails closed.
- `59d853618f` (review round): a plan whose view exceeds the plan's own
  limits declines to the from-scratch writer instead of failing the save.
- `f2a57ac936` (review round): the shared walkers are always inlined, which
  keeps the reader's instruction count at the base's, and the chain map is
  cleared by chain or by table, whichever is cheaper.

**Review round.** The independent review measured the first version
quadratic in mini streams: a Reuse write of 3,000 2,000-byte mini streams
went from 8.2 to 114 ms, and all-mini-stream fixtures regressed by a median
+7.4%. The four fixtures timed in round one hold 0, 1, 1 and 9 mini streams,
so round one never measured that path.
[The review-round section](#review-round-the-mini-stream-regression-and-its-fix)
has the fix and a three-arm re-measurement: base, first version, fix. With
the fix, every measured Reuse write is 3.9–40.9% faster than the base, the
readback's instructions grow ×3.07 from 1,000 to 3,000 mini streams, and
over the whole OLE2 corpus no fixture/edit pair runs more than 1.22% more
instructions than the base. The Result table below is round one's, measured
on the first version.
## Result

`OleWriter::write_to` with the default `SectorLayoutPolicy::Reuse` validates
its plan before emitting it. The validation's stream readback used to copy
every stream out of the planned artifact into a fresh buffer and compare it.
It now compares each stream where it lies. The structural reparse is
untouched, and on every plan tested the verdict is identical.

| Lane (8 paired layouts unless noted) | Fixture | Base p50 | Cand p50 | Paired Δ p50 [95% CI] |
|---|---|---:|---:|---:|
| CFB-only Reuse `write_to` | 45543.ppt | 31.9 µs | 22.0 µs | **−30.7%** [−34.4, −29.6] |
| | 41246-1.ppt | 25.5 µs | 17.0 µs | **−33.5%** [−34.7, −32.5] |
| | FloatingPictures.doc | 42.8 µs | 32.9 µs | **−23.5%** [−24.2, −21.3] |
| | NoHeadFoot.doc | 6.7 µs | 5.5 µs | **−17.1%** [−18.5, −15.3] |
| 0728 container control, Reuse (two writes per edit) | 45543.ppt | 164.3 µs | 139.1 µs | **−15.2%** [−15.5, −13.8] |
| | FloatingPictures.doc | 280.6 µs | 259.9 µs | **−7.3%** [−7.6, −6.2] |
| | NoHeadFoot.doc | 33.6 µs | 31.8 µs | **−6.1%** [−6.9, −5.2] |
| Public PPT slide removal (18 layouts) | 45543.ppt | 651.2 µs | 639.6 µs | −0.7% [−2.2, +3.2] |
| | 41246-1.ppt | 1,066.7 µs | 1,055.2 µs | **−1.05%** [−1.24, −0.71] |
| Public DOC paragraph replace (18 layouts) | FloatingPictures.doc | 1,050.2 µs | 1,051.3 µs | +0.05% [−1.1, +1.9] |
| | NoHeadFoot.doc | 77.8 µs | 72.8 µs | −6.1% [−8.7, −1.7] |

- **Per owner, the work removed is exact.** One Reuse write on 45543.ppt
  runs 137,122 fewer user instructions (−26.6%) and 50,464 fewer cycles
  (−34.4%). It allocates 386,906 fewer bytes in 28 fewer calls: the 384,906
  payload bytes the readback copied, plus its per-stream chain buffers.
  The container control removes exactly twice that and the public PPT
  removal exactly once.
- **Output bytes are unchanged.** Every process of both arms published the
  same output digest for its case, including the sealed 0728 digest
  `545ed7e5…` for the 45543.ppt removal.
- **Validation's share of the write falls from 52.7% to 19.9%** (Callgrind Ir,
  45543.ppt): 732.5 K to 163.4 K instructions per write. The structural
  reparse (84.3 K), planning (≈180 K) and emission (475.5 K) are unchanged.
- **The public lifecycles perform one Reuse write each,** worth about 1–3% of
  their instructions. Under the default glibc malloc, heap layout moves some
  of them by ±7–21% per round. The medians of the 45543.ppt removal, the
  FloatingPictures.doc replace and the harness DOC `large` save do not
  resolve the saving, and the removal's p95 is flagged.
- **With glibc's trim and mmap thresholds fixed** (a tested allocator-state
  mechanism), the 45543.ppt removal improves −3.50% [−4.19, −2.91] with no
  p95 flag. The DOC `large` save improves −1.96% [−4.12, −0.46].
- **Rewrite, read-only and no-op controls run the same instructions.** Their
  timings stay within noise, except the harness read selector
  `cfb_open/few-large` at +4.7%, a code-placement effect examined under
  [Regression flags](#regression-flags-every-change-above-5).

## Review round: the mini-stream regression and its fix

**The regression.** The independent review measured the first version
(`776ad05175`) against the base on the Reuse `write_to`:

| Input (review's measurement) | Base | First version |
|---|---:|---:|
| v3, 1,000 × 2,000-byte mini streams | 2.2 ms | 13.6 ms |
| v3, 3,000 × 2,000-byte mini streams | 8.2 ms | 114 ms |
| v4, 10,000 × 4,000-byte mini streams | 69 ms | 346 ms |

Over 418 real fixture/edit pairs, the 42 pairs on all-mini-stream files
regressed by a median +7.4%, 24 of them by more than 5%.

- **The mechanism.** For every mini stream, the first version's
  `minifat_stream_equals` repeated the root mini stream's whole load:
  - it re-collected the root chain;
  - it cleared a FAT-sized visited map;
  - it checked every root run for readability through `examine_at`.

  It then compared the stream one 64-byte mini sector at a time. The cost
  was mini streams × root sectors. `open_stream`, which it replaced, loads
  the mini stream once per reader. The verdicts were right; the cost was
  not.
- **Why round one missed it.** The fixtures it timed hold 0 (45543.ppt),
  1 (41246-1.ppt), 1 (NoHeadFoot.doc) and 9 (FloatingPictures.doc) mini
  streams, and its tests bound verdicts, not work per stream.
- **What the fix could not remove.** It is also an adversarial cost: a
  source crafted with many mini streams made validation quadratic. The fix
  removes that from the readback. The unchanged reparse's own per-stream
  term (below) predates this record.

**The fix** (`56548fd06e`, `59d853618f`, `f2a57ac936`):

- **`StreamComparer` binds the comparison to one reader and one candidate.**
  It performs the root load once, on the first mini or empty stream. It
  keeps the load only after it succeeded; a failed load is repeated and
  fails the same way, as `open_stream` does.
- **Mini sectors are compared as one range per run.** Sectors whose bytes
  follow one another in the file become one candidate call, instead of one
  call per 64-byte mini sector.
- **Chain maps are cleared in proportion to the work.** `EndChainScratch`
  restores its visited map by clearing the chain's bits or the table's
  words, whichever is fewer. The first version cleared a table-sized map
  per chain.
- **The review's nits are fixed.**
  - `examine_at` fails closed.
  - The reader and the comparison share `walk_chain_to_end` and
    `visit_sector_runs`, always inlined, so they cannot drift.
- **Plans whose view exceeds the plan's own limits decline.** This was a
  pre-existing gap the review found, identical on the base. The review saw
  it in 12 of its fault cases: a FAT that runs the version 3 directory chain
  past the planned directory image makes the reparse report `LimitExceeded`
  ("directory bytes"). `write_to` failed the save instead of falling back.
  - The validation reader's only limits are the plan's own output and
    directory lengths (`OleFileLimits::for_writer`), so this is a plan
    defect of the same class as the `CorruptedFile` errors that already
    decline.
  - `plan_validation_declines` now declines it. Allocation and
    invalid-limit failures still fail the save.
  - A named fault test covers both validators.

An intermediate build (`59d853618f`) showed two small costs of the
de-duplication:

- The reader's open-and-read-all of 45543.ppt ran 1,962 more instructions
  (159,634 → 161,596), because the shared walkers were no longer inlined.
- A Reuse write of 45543.ppt ran 5,531 more instructions than the first
  version, because per-bit clearing costs more than word clearing for long
  regular chains.

`f2a57ac936` fixes both: the reader runs 159,420, and the write 379,453.
That intermediate matrix is kept as `analysis-superseded-59d853618f.json`.

**Re-measurement.** Three arms, built with identical commands (rustc
1.95.0), in `round2/` of the packet:

| Arm | Commit | What it is |
|---|---|---|
| A | `ab29ac6291` | Base |
| B | `776ad05175` | First version |
| C | `f2a57ac936` | Fix |

- **Inputs:**
  - the four round-one fixtures with their public edits;
  - hyperlink.doc, `test-data/ole/ppt/empty.ppt` and WithCheckBoxes.xls,
    with both edits of the `sector_layout_corpus` test (`same` flips the
    largest stream's last byte; `grow` appends `3 * sector_size + 7` bytes);
  - four files generated by the probe (`generate`) with the `grow` edit:
    v3 1,000 and 3,000 × 2,000 B, v4 3,000 × 2,000 B and 10,000 × 4,000 B.
- **Controls:** the Rewrite policy, the reader, and the harness edit/save
  and read selectors (A and C only).
- **Timing:** 12 rounds, each of the six A/B/C orders twice, with argv[0]
  8 bytes longer each round, pinned to core 20. 660 processes, none failed,
  and every output digest is identical across all arms.
- **Further lanes:** per-owner `instructions` and `cycles` (three layouts),
  Callgrind scaling, a corpus-wide instruction lane and the allocation lane.

`summary-r2.md` has every row.

| Reuse `write_to` | A p50 | B p50 | C p50 | C vs A [95% CI] | B vs A | C vs B |
|---|---:|---:|---:|---:|---:|---:|
| v3 1,000 × 2,000 B | 1,843 µs | 15,579 µs | 1,719 µs | −6.44% [−7.59, −5.21] | +742% | −88.9% |
| v3 3,000 × 2,000 B | 7,149 µs | 125,726 µs | 6,428 µs | −10.55% [−12.31, −8.61] | +1,659% | −94.9% |
| v4 3,000 × 2,000 B | 7,077 µs | 22,377 µs | 6,384 µs | −10.47% [−11.36, −8.89] | +215% | −71.6% |
| v4 10,000 × 4,000 B | 63,218 µs | 377,096 µs | 46,340 µs | −26.12% [−26.67, −25.68] | +490% | −87.7% |
| hyperlink.doc, same | 6.8 µs | 7.5 µs | 6.5 µs | −4.79% [−5.37, −3.52] | +10.30% | −13.59% |
| hyperlink.doc, grow | 6.9 µs | 7.1 µs | 6.5 µs | −5.70% [−6.89, −4.54] | +2.68% | −7.95% |
| empty.ppt, same | 6.4 µs | 7.2 µs | 6.1 µs | −4.83% [−6.04, −4.52] | +11.85% | −15.25% |
| empty.ppt, grow | 6.6 µs | 6.9 µs | 6.4 µs | −3.85% [−4.35, −2.65] | +4.68% | −7.91% |
| WithCheckBoxes.xls, same | 13.5 µs | 15.3 µs | 12.5 µs | −7.03% [−7.55, −6.57] | +13.51% | −17.96% |
| WithCheckBoxes.xls, grow | 13.7 µs | 14.8 µs | 12.6 µs | −8.19% [−8.62, −7.68] | +8.32% | −14.82% |
| 45543.ppt, public | 38.5 µs | 22.8 µs | 22.7 µs | −40.94% [−41.83, −40.32] | −40.77% | −0.72% |
| 41246-1.ppt, public | 28.8 µs | 19.6 µs | 19.2 µs | −33.48% [−35.09, −32.03] | −32.03% | −1.92% |
| FloatingPictures.doc, public | 45.8 µs | 32.5 µs | 31.1 µs | −32.13% [−33.03, −31.28] | −28.90% | −4.59% |
| NoHeadFoot.doc, public | 6.5 µs | 5.3 µs | 5.4 µs | −16.88% [−17.19, −16.29] | −17.30% | +1.03% |

Controls, C vs A:

- Rewrite: 45543.ppt −3.71%, WithCheckBoxes.xls −0.55%.
- Reader: open-and-read-all of 45543.ppt +0.88% [−1.81, +3.54].
- Harness edit/save:
  - `doc_semantic_one_edit_save`: tiny −2.20%, large +2.83% [−1.91, +10.63];
  - `ppt_semantic_one_edit_save`: tiny −0.20%, large −1.19%;
  - `xls_semantic_one_edit_save`: tiny +0.43%, large −0.68%.
- Harness reads:
  - `cfb_open`: few-large +0.19%, many-small +0.20%;
  - `cfb_read_one`: few-large +1.68% [+0.87, +2.98], many-small +0.00%.

**Instructions per owner** (user mode, median of three layouts):

| Case | A | B | C | C vs A | C vs B |
|---|---:|---:|---:|---:|---:|
| v3 1,000 × 2,000 B | 28.13 M | 384.52 M | 25.96 M | −7.69% | −93.25% |
| v3 3,000 × 2,000 B | 87.42 M | 3,268.04 M | 79.63 M | −8.91% | −97.56% |
| v4 3,000 × 2,000 B | 82.83 M | 491.93 M | 75.96 M | −8.29% | −84.56% |
| v4 10,000 × 4,000 B | 457.29 M | 9,225.91 M | 417.78 M | −8.64% | −95.47% |
| hyperlink.doc, same | 127,502 | 151,278 | 123,409 | −3.21% | −18.42% |
| empty.ppt, same | 126,417 | 150,250 | 121,217 | −4.11% | −19.32% |
| WithCheckBoxes.xls, same | 261,461 | 324,547 | 250,158 | −4.32% | −22.92% |
| 45543.ppt, public | 518,694 | 383,706 | 379,453 | −26.84% | −1.11% |
| NoHeadFoot.doc, public | 112,055 | 96,198 | 95,916 | −14.40% | −0.29% |
| Reader, 45543.ppt | 159,608 | 159,608 | 159,420 | −0.12% | −0.12% |

Every harness selector runs −0.97% to +0.11% of the base's instructions.

**Scaling.** Callgrind per write on the generated v3 files:

| | Streams | `write_to` | `validate` | Reparse | Of which A5 | Readback |
|---|---:|---:|---:|---:|---:|---:|
| A | 1,000 | 38.18 M | 22.99 M | 9.77 M | 6.10 M | 12.39 M (`open_stream`) |
| A | 3,000 | 164.47 M | 116.99 M | 53.13 M | 42.30 M | 61.39 M |
| B | 1,000 | 394.08 M | 378.89 M | 9.76 M | 6.10 M | 368.86 M |
| C | 1,000 | 31.57 M | 16.38 M | 9.76 M | 6.10 M | 6.36 M (`StreamComparer`) |
| C | 3,000 | 120.92 M | 73.41 M | 53.08 M | 42.30 M | 19.53 M |

Callgrind counts its own byte-loop `memcpy` and `memset`, so its totals
exceed the native counts above; its ratios are what matter here.

- **C's readback is linear:** ×3.07 for three times the streams. The base's
  readback grows ×4.96, because it allocates a zeroed map sized to the whole
  MiniFAT for each stream.
- **The reparse is the one super-linear term left,** unchanged and identical
  in every arm. Its `validate_stream_allocations` (A5) clears a
  table-sized map per stream: ×6.93 from 1,000 to 3,000 streams.
- **Natively,** C's whole write runs ×3.07 the instructions (25.96 M →
  79.63 M) and the base's ×3.11.

A5 runs on every `OleFile::open` of the reader, not only here; see the
follow-ups.

**The corpus.** Every OLE2 fixture (214), with both corpus edits: 418
fixture/edit pairs, all taking Reuse, and every output identical across the
three arms. User instructions per owner:

| Pairs | B vs A, median | B > +5% | C vs A, median | C vs A, max | C > +1% | C > +5% |
|---|---:|---:|---:|---:|---:|---:|
| All 418 | −11.80% | 57 | −12.55% | +1.22% | 1 | 0 |
| 42 on all-mini-stream files | +14.06% | 33 | −3.87% | +1.22% | 1 | 0 |

- **This reproduces the review's finding in instructions.** B's all-mini
  median is +14.06%, where the review timed +7.4%; the fix removes it.
- **The one pair above +1% is WithExtendedStyles.xls, same-length edit:**
  +1.22% (88,631 → 89,715), where B was +16.36%. The same file's `grow`
  edit is −5.27%.

**Allocation.** B and C allocate identically in every case; the fix
allocates nothing new. Against the base, the v4 10,000-stream write
allocates 189 MB instead of 1,063 MB (121,771 calls instead of 201,761).
Peak live bytes are unchanged, except that WithCheckBoxes.xls `grow` falls
from 43,191 to 42,620.

**Regression flags in the re-measurement** (C vs A, a round's paired p50,
mean or p95 above +5%): 33, each listed in `summary-r2.md`.

- **`doc_semantic_one_edit_save/large`, 12 flags.** Rounds 3, 4, 5, 8 and 9
  are +6.2% to +27.5%.
  - Instructions are equal (−0.02%).
  - The candidate takes 478–605 page faults per owner where the base takes
    about 0: the glibc trim policy that round one tested.
  - With `glibc.malloc.trim_threshold` and `mmap_threshold` at 256 MiB,
    both arms take 0 page faults, and C runs +0.43% cycles and −0.23%
    instructions.
- **`doc_semantic_one_edit_save/tiny`, 5 p95 flags:** rounds 2, 3, 4, 6
  and 10, +5.7% to +15.7%. Its p50 is −2.20%.
- **`open-read-45543`, 5 flags:** round 2 p95 +9.4%, and rounds 8 and 10
  p50/p95 +5.3% to +6.0%. Its instructions are −0.12%.
- **`cfb_read_one/few-large`, 4 flags:** rounds 5 and 11, +5.4% to +5.8%,
  and +1.68% over all rounds. Instructions and cycles per owner are the
  same (−0.00% and −0.12%): code placement.
- **`cfb_read_one/many-small`, 4 flags:** rounds 8 and 11, 0.3 µs → 0.4 µs,
  up to +21.9%. Instructions +0.11%.
- **`rewrite-45543-public`, 2 flags:** round 3, p50 +6.3% and mean +5.5%.
  Instructions −0.15%.
- **`cfb_open/many-small`, 1 flag:** round 1 p95 +5.3%.

None of the 14 CFB-only Reuse write cases, where this change's work is, has
a flag.

## What was changed

The state after the review round.

`crates/litchi-cfb/src/file.rs` gains an in-place comparison beside the
reader:

- `CandidateBytes` (crate-private trait): `equals_at(offset, expected)` and
  `check_readable(offset, len)` over the bytes of a composed candidate.
- `StreamComparer` (crate-private), created by
  `OleFile::stream_comparer(&candidate)`. It borrows one reader and one
  candidate, so its cached state cannot outlive them or be applied to
  others.
  - `stream_equals(path, expected)` performs `open_stream`'s traversal, with
    each range it would read handed to the candidate instead of copied.
  - `load_root` performs the root mini stream's load once per comparer and
    keeps it only after it succeeded, as `open_stream` caches its loaded
    mini stream.
  - `minifat_stream_equals` compares mini sectors whose bytes follow one
    another in the file as one range.
- `EndChainScratch`: reusable chain buffers whose visited map is restored
  by clearing the chain's bits or the table's words, whichever is fewer. So
  collecting chains costs time proportional to the chains.
- `walk_chain_to_end`: the one implementation of `collect_sector_chain`'s
  loop, which `collect_sector_chain` and `EndChainScratch` both call.
- `visit_sector_runs`: `read_sectors_batched`'s run segmentation and checks,
  which `read_sectors_batched` itself now runs on. Both shared walkers are
  `#[inline(always)]`, so the reader keeps the base's instruction count.
- Helpers: `present_bytes` (the body of `present_sector_bytes`),
  `read_bytes_equal`, `ministream_physical_offset`, `file_bytes_equal`,
  `CheckedBitSet::remove`.

`crates/litchi-cfb/src/writer/layout.rs`:

- `ReusePlan::validate` keeps its structural reparse and now runs a
  `StreamComparer` over a `PlanView`, whose `examine_at` walks the planned
  bytes exactly as `ReusePlan::read_at` copies them. Every comparison goes
  through `fold_comparison`, which fails closed: an expectation that does
  not cover a range is a mismatch.
- The previous validation is kept as `#[cfg(test)] validate_by_readback`, the
  oracle the tests hold the new one to. `ReusePlan` derives `Clone` under
  `cfg(test)` for fault injection.

`crates/litchi-cfb/src/writer/core.rs`: `plan_validation_declines` becomes
`pub(super)` so the tests use the writer's own decline classification. It
now also declines `OleError::LimitExceeded`: the validation reader's only
limits are the plan's own output and directory lengths, so exceeding one is
a plan defect (see the review-round section).

New tests: `src/writer/layout/validation_tests.rs`,
`src/stream_compare_tests.rs`, and three `file::tests` for the chain
scratch.

There is no public API change, no new dependency, no `unsafe`, no new budget
or limit, and no change to emitted bytes, planning or emission.

## What the Reuse-plan validation proves

`OleWriter::write_to` plans a Reuse layout (`plan_reuse`), validates it
(`ReusePlan::validate`), and only then emits it (`ReusePlan::emit`). Emission
writes every output sector from the plan in ascending order. Validation sees
the same plan through `PlanCursor`, a positional view (`ReusePlan::read_at`).
The view composes the header, the rebuilt FAT, MiniFAT, directory and
mini-stream images, and the model's payload chunks sector by sector, without
building the artifact. A declining validation error makes the writer fall
back to the from-scratch serialization (`SectorLayoutFallback::PlanRejected`).
Nothing reaches the sink first.

Validation has two parts. This change keeps the first byte for byte and
reaches the second's verdict by a different route.

**A. Structural reparse** (`OleFile::open_with_limits` over the view, under
`OleFileLimits::for_writer`; unchanged):

| # | What is proven | Where |
|---|---|---|
| A1 | Header: magic, byte order, zero reserved bytes, version/sector-shift pairing, the directory-sector count rule per version, mini-sector shift 6 and cutoff 4096, MiniFAT and DIFAT start/count agreement, a directory start below `MAXREGSECT`, every declared count within the physical sector count, and a length within the writer limit | `open_with_limits` |
| A2 | FAT: the header lists exactly the declared FAT sectors, each in bounds and owned once, and each is marked `FATSECT` in the table it holds | `load_fat`, `claim_sector` |
| A3 | Directory: the chain is acyclic, in bounds, of the declared (v4) or limited (v3) length and owned once. Every entry has a known type and colour, a legal NUL-terminated UTF-16 name, and valid storage or stream fields, with the root at SID 0. The sibling trees are strictly ordered by MS-CFB name comparison and acyclic, and every allocated SID is owned by exactly one storage | `load_directory`, `validated_directory_entries` |
| A4 | MiniFAT: exactly the declared count, acyclic, in bounds, owned once | `load_minifat` |
| A5 | Allocation: the root mini-stream chain is exact for the root size. Every stream's chain is exact for its declared size: `ceil(size/64)` mini sectors inside the root mini stream and not shared with another stream, or `ceil(size/sector)` regular sectors owned once. Exact means a regular start sector (`ENDOFCHAIN` for an empty chain), no cycle, and `ENDOFCHAIN` after exactly that many links | `validate_stream_allocations` |
| A6 | Partition: every physical sector that no role owns is `FREESECT`, and the FAT covers every physical sector | `validate_physical_sector_layout` |

**B. Readback.** For every model stream, the reader must find its path in
the validated directory by MS-CFB name comparison, the entry must be a
stream, and reading it must yield exactly the model's bytes. A regular
stream is read at its declared size through its FAT chain. A mini stream is
read through its MiniFAT chain inside the root mini stream, which is loaded
through the FAT. B therefore proves:

- **B1:** each model path exists as a stream entry.
- **B2:** its declared size equals the model length.
- **B3:** chunk k of the payload is what the k-th sector of its chain holds,
  for regular and mini streams alike. So the planner's path-to-SID mapping,
  its chunk order and its mini-stream image are right.
- **B4:** every range the reader touches is readable in the view. That
  includes the whole root mini stream once any mini or empty stream is read.

It does **not** prove, and never did:

- **That unchanged streams kept their source sectors.** That is a placement
  policy, counted by `SectorLayoutReport` and tested by
  `sector_layout_smoke`.
- **That directory metadata outside the planner-owned fields survives.**
  Construction ensures it, and the 0663 and 0728 normalized-directory checks
  test it.
- **That the directory holds no stream beyond the model's.** The planner's
  shape gate requires the stream and storage sets to match exactly.
- **That `emit` writes what `read_at` shows.** Both read one plan. The corpus
  test reopens every emitted output and compares it stream by stream with the
  Rewrite output.

**What depends on each part.**

- **A1–A6** are the ordinary reader's admission rules, and every consumer
  reopens a published OLE2 artifact through them:
  - the common editor's candidate commit reopens and recaptures the
    rendering;
  - PPT finish reopens it to check the persist mapping (`validate_rewrite`);
  - every public snapshot reopens its bytes;
  - a later save adopts the output as its next source, and
    `SourceLayout::parse` requires the same partition.

  ADR 0006 (malformed output is refused before publication, and normal save
  never repairs) and ADR 0005 (a sink never observes an unvalidated
  candidate) rest on validation finishing before `emit`.
- **A5 and A6** stop one stream aliasing another's bytes, and stop a released
  sector's stale bytes leaking through a shared or orphaned sector.
- **B2** is what every directory-size consumer trusts: `stream_len`,
  `read_stream_range`, and the format parsers that size their reads from it.
- **B1 and B3** are the logical publication itself. The committed snapshot's
  bytes are the patch's after-image (ADR 0003), and untouched streams must be
  byte-identical (ADR 0006 preservation). A wrong SID mapping or chunk order
  would publish a valid container with the wrong bytes at a path; only B
  catches it.
- **B4** keeps a defect in unused mini-stream space from passing.

## The same proof, cheaper

**Where the time went.** A native frame-pointer profile of the Reuse
`write_to` on 45543.ppt (20,000 owners, restricted to the timed owner) gives
this split. The reparse is unchanged by this record; planning and emission
are outside it.

| Phase | Base | Candidate |
|---|---:|---:|
| Plan | 32.6% | 45.5% |
| Stream readback (per-stream `open_stream` and compare) | 29.4% | 12.9% |
| Reparse (A1–A6) | 11.8% | 16.3% |
| Emission | 4.2% | 4.6% |
| Unattributed leaf `memcmp`/`memmove` (callers lost to frame-pointer unwinding) | 20.9% | 20.0% |

The shares are of different totals: the candidate's `write_to` takes about
0.69 of the base's time. The base's unattributed leaves include the
readback's whole-stream `memcmp` (10.4% on its own). Callgrind, with exact
inclusive attribution, puts the base readback (`open_stream` and `memcmp`) at
645.7 K of 1,389.7 K Ir per write. The candidate's `stream_equals` takes
77.1 K Ir, and validation as a whole goes from 732.5 K to 163.4 K Ir per
write. Per stream, the old readback did the following:

- allocated a zero-filled buffer of the stream's declared size;
- walked the chain into a fresh vector with a fresh visited map sized to the
  whole FAT;
- read every sector through `PlanCursor`, which cleared each sector and then
  copied the payload chunk into it;
- compared the whole buffer with the model.

That is four passes over every byte and three allocations per stream. For a
correct plan, the bytes it copied are the model's own payload chunks.

**What the new readback does.** `StreamComparer::stream_equals` performs
`open_stream`'s traversal in the same order and with the same errors:

1. the entry lookup and stream-type check;
2. the chain collection (`EndChainScratch::collect`, which runs
   `collect_sector_chain`'s own walker: index in the table, no revisit, only
   `ENDOFCHAIN` or regular links);
3. the declared-length and chain-capacity checks;
4. for a mini or empty stream, the root mini-stream load checks, once per
   comparer as `open_stream` loads the mini stream once per reader;
5. the mini-sector bounds;
6. `read_sectors_batched`'s run segmentation and physical bounds checks
   (`visit_sector_runs`, which the batched read itself runs on).

Where `open_stream` would copy a physical range into its result,
`stream_equals` asks the candidate instead. It compares the range while the
stream still matches and only checks it for readability once it does not. So
every read error the old path raised is still raised, including one in root
mini-stream space no mini stream uses. The zero fill of a short final sector
is compared with zero, and a length mismatch still visits every range.

**`PlanView::examine_at` compares a range where `read_at` would copy it.** It
walks header bytes, table images, unallocated sectors and payload chunks. It
joins consecutive sectors whose source continues (`continues`: the next
position of the same table image or the next chunk of the same stream, with
checked successors) and examines each run as one slice. Its bounds and
overflow checks fail exactly when one of the run's sectors' would. A payload
slice that is the very slice of the model it is compared with, the same
address and length, is equal without being read: equality is reflexive, and
both are shared borrows for the whole validation. Every other slice is
compared byte for byte:

- a chunk planned from the wrong stream or at the wrong position;
- the mini-stream image, which the planner copied;
- table images;
- zero padding.

**Why the verdict cannot differ.**

- The structural reparse is the same call on the same view.
- For each stream, `stream_equals` reaches the same outcome as
  `open_stream` followed by a comparison. It returns `true`, `false` or an
  error exactly when that sequence would, because:
  - its traversal is `open_stream`'s, with each check and error string;
  - the bytes it examines are the bytes `read_at` would copy, from the same
    sources;
  - a skipped `memcmp` compares a slice with itself.
- The root load's result is kept exactly as `open_stream` keeps its loaded
  mini stream: once, after a successful load, for the one reader. A failed
  load is repeated for the next mini stream and fails the same way.
- Comparing contiguous mini sectors as one range examines the same bytes:
  the zero fill of a short final sector depends only on the physical
  position, which a contiguous range preserves. Those comparisons cannot
  raise an error the load did not, because the load already checked every
  range of the root chain for readability.
- Two differences remain, and neither changes the verdict:
  - An allocation failure of a removed buffer can no longer occur.
  - A read error is the view's `InvalidData` rather than the same failure
    wrapped by `PlanCursor` in `OleError::Io`. The writer declines both
    alike (`plan_validation_declines`).

**Not done.** Three shortcuts were considered and rejected:

- **Trusting unchanged or unmoved streams.** Every sector of every stream is
  still located by the reader's own validated tables and checked against its
  planned source. The identity shortcut only applies where that check has
  already shown the source is the model's own slice.
- **Fusing validation with emission.** A sink must not observe an
  unvalidated candidate, so validation stays a complete pass before
  `emit`.
- **Dropping the reparse or the readback's chain walks.** A5 already walked
  every stream chain over the same tables, so the readback's walk is
  redundant with it. It is part of the candidate readback's 2.8 µs (12.9% of
  a 22 µs 45543.ppt write in native samples) and is kept rather than argued
  away.

## Authority and constraints

- **ADR 0006:** preservation and validation. Every check of A1–A6 and B1–B4
  still runs before a byte is emitted, and validation still never mutates.
  Refusals are unchanged: declining plans still fall back to the from-scratch
  writer. Output bytes are identical in every measured process.
- **ADR 0005:** a sink still never observes an unvalidated candidate.
  Validation's transient memory falls by the size of every stream it
  compared; peak live bytes of the measured owners are unchanged, because
  emission sets that peak.
- **ADR 0003:** committed bytes, and therefore patches and their inverses,
  are unchanged.
- **ADR 0026:** directory metadata validation stays in `litchi-cfb`'s reader,
  unchanged.
- **Change 0652, decision 10** made Reuse the default. This record lowers the
  cost of that default without changing its layouts or its Rewrite fallback.
- **Change 0652's standing trade-offs:**
  - Correctness comes first: the verdict is proven the same, and tested on
    5,600 random faults.
  - Optimize the benign common path: a correctly placed payload run
    compares by identity, while a corrupted plan pays for byte comparison.
- **GOAL, "LEGACY CFB-SPECIFIC WORK":** all CFB validation, ownership, cycle,
  overlap, FAT, MiniFAT, directory and truncation checks are preserved. No
  ODF or iWork crate is touched, and record 0748's `overlay.rs` is untouched.

## Proof obligations and tests

- **Named faults** (`validation_tests.rs`). Every case must get the same
  verdict from the in-place validation and from the retained readback oracle,
  and a non-declining error fails the test.
  - Refused by both:
    - one chain running into, or joining, another;
    - two entries sharing a chain;
    - self and back cycles;
    - a chain ending early or linking `FREESECT`;
    - a declared size ±1 for regular, mini and nested streams;
    - a model one byte longer or shorter;
    - chunks swapped between streams or reordered within one;
    - a shifted chunk index;
    - a payload sector left unallocated;
    - a missing stream index;
    - a changed mini-stream payload byte;
    - overlapping or cyclic mini chains;
    - a truncated mini-stream image (only the root mini-stream load reaches
      it);
    - a moved directory start;
    - a cleared MiniFAT count;
    - an unmarked FAT sector;
    - a truncated FAT image;
    - a model byte changed after planning in a mini stream.
  - Accepted by both:
    - a flipped byte in mini-stream padding past a stream's end;
    - swapped chunks with identical bytes;
    - two paths sharing one payload allocation, with their chunks crossed;
    - a model byte changed after planning in a regular stream. A regular
      stream's sectors are planned as chunks of whatever payload the model
      supplies, so readback and emission see the same bytes; only its
      length is baked into the plan.
- **Random faults.** Bit flips in the header, FAT, MiniFAT, directory and
  mini-stream images, swapped, freed or re-planned sectors, and rewritten FAT
  entries give identical verdicts in all 5,000 cases:

  | Plan | Faults | Accepted by both | Refused by both |
  |---|---:|---:|---:|
  | Synthetic, 512-byte sectors | 2,000 | 839 | 1,161 |
  | Synthetic, version 4 | 1,000 | 642 | 358 |
  | 45543.ppt | 500 | 185 | 315 |
  | 41246-1.ppt | 500 | 204 | 296 |
  | FloatingPictures.doc | 500 | 132 | 368 |
  | NoHeadFoot.doc | 500 | 284 | 216 |

- **`stream_equals` against `open_stream`** (`stream_compare_tests.rs`).
  Results are identical, including every error string:
  - every stream of 214 OLE2 fixtures, with seven expectations each (exact,
    three single-byte flips, one byte longer, one shorter, empty): 9,041
    comparisons;
  - 2,884 of 3,000 corrupted files that still open;
  - files with short final sectors, in both geometries.
- **Chain collection:** `EndChainScratch::collect` matches
  `collect_sector_chain`'s result or error on 14 adversarial tables. Its
  visited map is clear after every call, success or error, and a long chain
  followed by short ones through a 10,000-entry table reuses both buffers.
- **Review round, linearity and caching**
  (`stream_compare_tests.rs`):
  - over 300 mini streams in both geometries, the root load's readability
    checks happen for the first mini stream only, and each contiguous stream
    takes exactly one candidate comparison (not one per 64-byte mini sector);
  - a root load that fails once is not kept: the next mini stream loads
    again and succeeds, and a load that always fails fails every mini stream
    the same way;
  - files of 120 mini streams in both geometries reach the readback verdict,
    call for call.

  Sequence equivalence is now checked with a fresh reader and a fresh
  comparer seeing the same calls in the same order, so each side's one-time
  mini-stream load is exercised alike.
- **Review round, plans** (`validation_tests.rs`):
  - 240-mini-stream plans in both geometries, with mini growth, a shrink
    and a mini-to-regular migration, are accepted by both validators;
  - 600 random faults on them give identical verdicts (512-byte: 60
    accepted, 240 refused; 4096-byte: 128 accepted, 172 refused);
  - the directory-chain-past-its-image fault raises `LimitExceeded` in both
    validators and now declines;
  - the classifier keeps allocation and invalid-limit failures on the error
    path.
- **Admission is unchanged.** The existing corpus test
  `every_ole2_fixture_agrees_stream_by_stream_under_both_policies` admits
  Reuse for 209 of 211 examined fixtures on no-op, same-length and
  length-changing saves, 195 of them appending. These are 0663's counts.

## Method

- **Base and candidate** are `ab29ac6291` and `897d24fd0f`, built with
  identical cargo commands, flags and features under rustc 1.95.0.
  `binaries.json` in the packet has the commands and every SHA-256:

  | Binary | Base | Candidate |
  |---|---|---|
  | probe | `d1f026b9…6745` | `0646fdbf…0f24` |
  | probe_alloc | `392bfa8b…5e45` | `75b4323d…0075` |
  | litchi-perf-baseline | `773f050f…9978` | `77b88937…794f` |

  A repeated harness build compiled nothing in either arm.
- **The probe** (`probe/`, public APIs only) has these modes:
  - `ppt-remove`, the 0728/0734/0745 public slide removal (open, edit,
    remove slide 1, commit, output copy);
  - `doc-replace`, the 0728/0745 public paragraph replace;
  - `container-{reuse,rewrite}`, the 0728 common-container control, which
    stages the exact streams an untimed public edit changed and finishes;
  - `cfb-write-{reuse,rewrite}`, one `OleWriter::write_to` of the edited
    model over the adopted original;
  - `ppt-noop` and `cfb-open-read`, controls.

  Harness selectors: `doc/ppt/xls_semantic_one_edit_save` and
  `…noop_edit_save` with writer shapes `tiny` and `large`, and `cfb_open`
  and `cfb_read_one` with shapes `few-large` and `many-small`. A sampled
  profile shows the DOC and PPT saves reach `plan_reuse` and `examine_at`.
  `xls_semantic_one_edit_save` is a same-length edit and takes the
  copy-through overlay instead, so it is a control here.
- **Timing.** Every process is pinned with `taskset -c 20` on the 32-core
  AMD EPYC 9R45, with warm caches, while other agents build on other cores.
  - Eight rounds, in ABBA order across rounds, with case order rotated. Each
    round starts every binary through a symlink 8 bytes longer, 0745's
    heap-layout randomization.
  - The public selectors, the harness edit/save and CFB read selectors and
    the flagged Rewrite control were continued to 18 rounds.
  - Samples per process: 60 for the public lifecycles, 200–1,000 for the
    container, 2,000–5,000 for the CFB-only write, and 100 or 200 for the
    harness, after 5–100 warmups.
  - 492 processes, none failed.
  - Statistics: process p50 is the midpoint median and p95 the nearest rank.
    Each comparison is the median of the per-round paired change; its
    interval is a percentile bootstrap of that median (10,000 resamples,
    seed 749).
- **Counters.** Per-owner `instructions`, `cycles`, their `:u` forms and
  page faults, from two processes of 20/220, 50/550 or 200/2,200 owners at
  three layouts. Probe processes run with `--oracle first`, so the
  difference holds only the timed owner and its loop.
  - A first pass, kept as `per-owner-v1-digest-diluted`, digested every
    owner's output; that SHA-256 diluted its percentages.
  - Supplementary lanes: the harness read controls (40/240 owners), and
    the DOC edit/save across ten layouts with and without fixed glibc
    malloc thresholds.
- **Allocation.** One counting-allocator process per case and arm, five
  owners each. Every owner was identical.

## Results

Selected rows; `summary.md` in the packet has every row, the counters for
every case and layout, and every flag.

| Case | Rounds | Base p50 µs | Cand p50 µs | Δ p50 [95% CI] | Δ mean | Δ p95 |
|---|---:|---:|---:|---:|---:|---:|
| `cfb-write-reuse-45543` | 8 | 31.9 | 22.0 | −30.73% [−34.35, −29.56] | −30.96% | −31.43% |
| `cfb-write-reuse-41246` | 8 | 25.5 | 17.0 | −33.53% [−34.67, −32.48] | −33.33% | −33.52% |
| `cfb-write-reuse-floating` | 8 | 42.8 | 32.9 | −23.45% [−24.17, −21.31] | −23.40% | −25.46% |
| `cfb-write-reuse-nohf` | 8 | 6.7 | 5.5 | −17.11% [−18.55, −15.28] | −17.25% | −18.02% |
| `cfb-write-rewrite-45543` | 18 | 6.6 | 6.8 | +3.07% [−2.45, +7.38] | +2.54% | +0.88% |
| `cfb-write-rewrite-floating` | 8 | 16.9 | 17.1 | +0.66% [−1.72, +2.84] | +0.84% | +1.06% |
| `cfb-write-rewrite-nohf` | 8 | 2.8 | 2.7 | −0.90% [−1.98, −0.36] | −0.93% | −0.73% |
| `container-reuse-45543` | 8 | 164.3 | 139.1 | −15.19% [−15.47, −13.80] | −15.10% | −14.82% |
| `container-reuse-floating` | 8 | 280.6 | 259.9 | −7.31% [−7.61, −6.21] | −7.35% | −7.11% |
| `container-reuse-nohf` | 8 | 33.6 | 31.8 | −6.05% [−6.90, −5.15] | −5.74% | −5.27% |
| `container-rewrite-45543` | 8 | 110.5 | 108.5 | −1.36% [−2.62, −0.06] | −1.42% | −1.19% |
| `container-rewrite-floating` | 8 | 149.6 | 149.2 | −0.09% [−0.91, +0.13] | −0.07% | −0.21% |
| `container-rewrite-nohf` | 8 | 24.5 | 25.0 | +1.31% [+0.30, +2.42] | +1.99% | +0.60% |
| `ppt-remove-45543` | 18 | 651.2 | 639.6 | −0.70% [−2.19, +3.21] | −0.06% | +6.96% |
| `ppt-remove-41246` | 18 | 1,066.7 | 1,055.2 | −1.05% [−1.24, −0.71] | −1.10% | −1.24% |
| `doc-replace-floating` | 18 | 1,050.2 | 1,051.3 | +0.05% [−1.10, +1.87] | −0.76% | −0.63% |
| `doc-replace-nohf` | 18 | 77.8 | 72.8 | −6.10% [−8.70, −1.72] | −5.40% | −5.44% |
| `doc_semantic_one_edit_save/tiny` | 18 | 37.2 | 36.3 | −2.14% [−2.63, −1.87] | −1.56% | −0.58% |
| `doc_semantic_one_edit_save/large` | 18 | 657.3 | 651.7 | +1.99% [−3.71, +5.76] | +3.20% | +1.25% |
| `ppt_semantic_one_edit_save/tiny` | 18 | 80.0 | 79.8 | −0.27% [−0.51, +0.26] | −0.23% | +0.67% |
| `ppt_semantic_one_edit_save/large` | 18 | 181.8 | 180.9 | −0.55% [−1.10, −0.24] | −0.65% | −0.69% |
| `xls_semantic_one_edit_save/tiny` (overlay path) | 18 | 80.8 | 80.3 | −0.33% [−0.88, −0.08] | −0.72% | −0.42% |
| `xls_semantic_one_edit_save/large` (overlay path) | 18 | 3,477.7 | 3,500.0 | +0.92% [−0.53, +1.36] | +0.82% | +0.48% |
| `ppt-noop-45543` | 8 | 57.9 | 57.9 | −0.20% [−1.10, +0.43] | −0.24% | −1.17% |
| `cfb-open-read-45543` | 8 | 10.7 | 10.9 | +1.57% [−1.39, +2.43] | +1.46% | +1.33% |
| `cfb_open/few-large` | 18 | 69.7 | 73.0 | +4.73% [+4.62, +5.25] | +4.77% | +4.66% |
| `cfb_open/many-small` | 18 | 76.5 | 75.2 | −1.57% [−1.81, −0.98] | −1.62% | −1.55% |
| `cfb_read_one/few-large` | 18 | 115.3 | 115.4 | −0.52% [−2.39, +1.08] | −0.28% | −0.00% |

**Per-owner counters** (median of three layouts; `summary.md` has each
layout):

| Case | instructions:u base → cand | Δ | cycles base → cand | Δ |
|---|---:|---:|---:|---:|
| `cfb-write-reuse-45543` | 515,648 → 378,526 | −26.6% | 146,664 → 96,200 | −34.4% |
| `cfb-write-reuse-41246` | 407,577 → 308,595 | −24.3% | 110,542 → 78,994 | −28.5% |
| `cfb-write-reuse-floating` | 604,959 → 506,865 | −16.2% | 186,022 → 144,920 | −22.1% |
| `cfb-write-reuse-nohf` | 112,823 → 97,017 | −14.0% | 29,404 → 25,203 | −14.3% |
| `container-reuse-45543` | 1,863,732 → 1,590,324 | −14.7% | 745,016 → 619,897 | −16.8% |
| `ppt-remove-45543` | 4,686,446 → 4,552,904 | −2.85% | 2,998,036 → 2,940,473 | −1.9% |
| `ppt-remove-41246` | 11,168,478 → 11,068,526 | −0.90% | 4,835,176 → 4,793,518 | −0.9% |
| `doc-replace-floating` | 10,957,271 → 10,858,866 | −0.90% | 4,760,046 → 4,780,708 | +0.4% |
| `doc-replace-nohf` | 1,020,892 → 1,011,456 | −0.92% | 318,431 → 338,660 | +6.4% |
| `doc_semantic_one_edit_save/large` | 41,408,749 → 41,383,476 | −0.06% | 10,404,274 → 11,833,050 | +13.7% |
| Every Rewrite control, `ppt-noop`, `cfb-open-read` | — | −0.51% to +0.31% | — | −3.1% to +0.9% |

The public PPT removals and the FloatingPictures.doc replace each save one
CFB-only write's worth: 133.5 K, 100.0 K and 98.4 K user instructions, and
exactly one write's bytes and calls in the allocation lane. So one Reuse
validation runs per public edit on these routes.

**Allocation** per owner: the Reuse write saves exactly the bytes the
readback allocated, and every other case is identical.

| Case | Bytes base → cand | Calls |
|---|---:|---:|
| `cfb-write-reuse-45543` | 1,635,017 → 1,248,111 (−386,906) | 143 → 115 |
| `cfb-write-reuse-41246` | 894,798 → 612,501 (−282,297) | 149 → 123 |
| `cfb-write-reuse-floating` | 1,346,403 → 1,003,359 (−343,044) | 308 → 240 |
| `cfb-write-reuse-nohf` | 110,654 → 85,829 (−24,825) | 132 → 105 |
| `container-reuse-45543` | 5,702,423 → 4,928,611 (−773,812) | 747 → 691 |
| `ppt-remove-45543` | 9,041,924 → 8,655,018 (−386,906) | 5,447 → 5,419 |
| `doc-replace-floating` | 14,115,027 → 13,771,983 (−343,044) | 15,141 → 15,073 |
| Every Rewrite control, `ppt-noop`, `cfb-open-read` | unchanged | unchanged |

Peak live bytes are unchanged in every case: emission's output buffer, not
the readback, sets the peak.

## Regression flags (every change above +5%)

The matrix has 116 per-round flags (a round's paired change of p50, mean or
p95 above +5%). `summary.md` lists each with its base and candidate values.
By case:

- **`cfb_open/few-large`**, a harness read selector: +4.73% [+4.62, +5.25]
  median p50, with 20 flags from +5.0% to +6.1% in 8 of 18 rounds.
  - No code on this path changed.
  - User instructions per owner are the same (+0.001% in the read-control
    lane).
  - Front-end counters at one layout show where the time goes: per owner the
    candidate takes about 35% more L1 instruction-cache fills from L2 (1,243
    against 923) and 49% more branch misses (134 against 90).
  - The probe's reader control, which runs the same `OleFile::open` and
    `open_stream` path in a different binary, has identical instructions and
    +1.57% [−1.39, +2.43].
  - This is code placement in the harness binary, whose `litchi-cfb` code
    moved when the new functions were added. It is kept as a flagged
    limitation, not claimed away.
- **`ppt-remove-45543`**: median p50 −0.70% and mean −0.06%, but **p95
  +6.96%**. 17 flags: p95 in 12 of 18 rounds (+5.1% to +9.2%), and p50 and
  mean in rounds 3, 4 and 12 (+6.1% to +7.7%).
  - The lifecycle takes about 440 page faults per owner in both arms, and its
    per-round p50 swings from −3.6% to +7.6% with layout.
  - Per owner it runs 2.85% fewer instructions and 1.9% fewer cycles.
  - The tail is not explained here; see the allocator-state finding below.
- **`doc_semantic_one_edit_save/large`**, 18 flags. Rounds 4, 5, 14 and 15
  have p50 +16% to +21%; other rounds are −16% to +6%.
  - The counters lane shows +13.7% cycles per owner, at −0.06% user
    instructions. In ten layouts the candidate takes 478–868 page faults per
    owner where the base takes −23 to 93; the tuned-glibc test below
    identifies the cause.
- **`cfb-write-rewrite-45543`**, a 6.6 µs Rewrite control, 18 flags.
  - Rounds 2, 3, 5, 9, 11, 12 and 13 have p50 from +5.7% to +16.6%; other
    rounds are −10.9% to +4.5%, for +3.07% [−2.45, +7.38].
  - Instructions (−0.26%) and cycles (−0.32%) per owner are the same.
- **`doc-replace-floating`**, 10 flags. Rounds 0, 2, 12 and 15 have p50 from
  +6.7% to +10.2%; the median is +0.05% [−1.10, +1.87].
- **`doc-replace-nohf`**, 7 flags. Round 1's mean is +61.1% from a few slow
  samples (its p50 is not flagged), and rounds 13 and 14 are +6.2% to
  +7.9%. The median p50 is −6.10%.
- **`doc_semantic_one_edit_save/tiny`**, 5 flags. p95 in rounds 0, 7 and 9
  (+7.0% to +18.6%), and round 16's mean (+87.2%) and p95 (+32.2%); its
  median p50 is −2.14% [−2.63, −1.87].
- **`cfb_read_one/many-small`**, a 0.3 µs read, 10 flags. They come from
  ±10 ns clock steps: 330 → 380 ns in round 1.
- **`xls_semantic_noop_edit_save/large`**, a 2 µs no-op, 4 flags: +28.8%
  p50 in round 0.
- **Single flags:**
  - `cfb_open/many-small`: round 12 p95 +11.3%;
  - `cfb_read_one/few-large`: round 4 p50 +7.6% and mean +5.5%;
  - `container-rewrite-nohf`: round 1 p95 +9.7%;
  - `doc_semantic_noop_edit_save/large`: round 0 p95 +7.5%;
  - `xls_semantic_one_edit_save/large`: rounds 3 and 14 p95 +7.1% and
    +5.6%.

**Allocator state, tested.** 0745 inferred that glibc's heap-trim and mmap
thresholds move these lifecycles with the early heap layout. The candidate
removes stream-sized temporary buffers, which changes where the surrounding
process's heap is trimmed and refaulted.

| `doc_semantic_one_edit_save/large`, per owner, 10 layouts | Page faults, base | Page faults, cand | Cycles, base → cand |
|---|---|---|---:|
| Default glibc | −23 to 93 | 478 to 868 | 10.40 M → 11.83 M (+13.7%) |
| `GLIBC_TUNABLES=glibc.malloc.trim_threshold=268435456:glibc.malloc.mmap_threshold=268435456` | 0 in every layout | 0 in every layout | 10.56 M → 10.33 M (−2.2%) |

- With the thresholds fixed, both arms stop faulting. The candidate is then
  2.2% cheaper per owner, with 0.17% fewer user instructions.
- In a direct, unsymlinked run of the same selector (whole process, 225
  iterations), the base took more than twice the candidate's page faults.
  `perf record` processed 132,364 page-fault events against 52,494, and kept
  71,640 against 32,839 after losing 45.9% and 37.4% while writing call
  chains. The direction depends on the process's early layout, not on the
  change.
- The mechanism is therefore glibc's trim policy reacting to the removed
  buffers. It is not added work, but on hosts with the default policy it can
  cost a process more page faults than the old buffers did, as it did in
  these layouts.

**Timing under fixed thresholds.** Eight more paired rounds of the public
and harness edit/save selectors ran with the same `GLIBC_TUNABLES` (packet
`tuned/`, 64 processes, identical output digests):

| Case | Base p50 µs | Cand p50 µs | Δ p50 [95% CI] | Δ mean | Δ p95 |
|---|---:|---:|---:|---:|---:|
| `ppt-remove-45543` | 419.8 | 406.1 | −3.50% [−4.19, −2.91] | −3.31% | −3.14% |
| `ppt-remove-41246` | 906.8 | 896.3 | −1.15% [−1.28, −0.67] | −1.00% | −1.27% |
| `doc-replace-floating` | 850.2 | 842.6 | −0.66% [−1.09, +0.13] | −0.85% | −1.14% |
| `doc_semantic_one_edit_save/large` | 655.7 | 637.6 | −1.96% [−4.12, −0.46] | −3.76% | −3.64% |
| `doc_semantic_one_edit_save/tiny` | 37.1 | 36.5 | −1.61% [−2.74, +0.30] | −2.65% | −6.80% |
| `ppt_semantic_one_edit_save/large` | 182.8 | 181.2 | −1.17% [−1.49, −0.22] | −1.17% | −1.64% |
| `ppt_semantic_one_edit_save/tiny` | 81.0 | 80.4 | −0.55% [−1.75, −0.18] | −0.46% | +1.63% |
| `xls_semantic_one_edit_save/large` (overlay path) | 3,437.6 | 3,436.7 | −0.76% [−1.31, +1.59] | −0.82% | +0.07% |
| `xls_semantic_one_edit_save/tiny` (overlay path) | 80.3 | 80.7 | +0.29% [+0.02, +0.67] | +0.26% | +0.92% |

- **With the allocator policy fixed, the 45543.ppt removal's p95 flag is
  gone.** Its p50 improves in every round (−1.6% to −5.1%), and the DOC
  `large` save improves instead of swinging with layout.
- **The tuned lane still has 15 per-round flags,** each a single-round mean
  or p95 outlier; `summary.md` lists them:
  - `doc-replace-floating`: round 2 mean +20.8% and p95 +71.0%;
  - `doc_semantic_one_edit_save/tiny`: round 1 mean +165.7% and p95 +15.8%;
    round 2 mean +38.2% and p95 +31.1%;
  - `ppt-remove-41246`: round 2 mean +7.7% and p95 +22.1%;
  - `ppt-remove-45543`: round 2 p95 +6.4%;
  - `ppt_semantic_one_edit_save/tiny`: round 0 mean +79.8% and p95 +10.0%;
    round 3 mean +187.9% and p95 +7.6%;
  - `xls_semantic_one_edit_save/tiny`: round 3 mean +41.6%; round 4 p95
    +5.7%.

  Round 2 flags four cases at once, which points at a host disturbance during
  that round. Outliers of the same size also go the other way: the 45543.ppt
  removal's round 1 p95 is −71.3%.
- **Fixing the thresholds also makes the base 45543.ppt removal 35% faster
  than under the default policy** (651.2 → 419.8 µs p50). That is an allocator
  finding, outside this change's scope.

## What is not claimed

- No registered speedup: `performance_claim: none`.
- **No end-to-end speedup for the public 45543.ppt removal or the
  FloatingPictures.doc replace.** Under the default allocator policy, their
  medians do not resolve the 1–3% instruction saving under ±7–21% layout
  swings, and the removal's p95 is flagged above. The fixed-threshold
  improvements are evidence about mechanism, not a claim for default
  deployments.
- **No claim that validation is now cheap in general.** The structural reparse
  (16% of the candidate's CFB-only write) and the readback's chain walks
  remain, by design.
- **No claim that validation as a whole is linear in the number of
  streams.** The readback is; the unchanged reparse's A5 is not (see the
  follow-ups).
- **No claim for every file.** One of 418 corpus fixture/edit pairs
  (WithExtendedStyles.xls, same-length edit) runs 1.22% more instructions
  than the base.
- **No claim about planning or emission.** Planning is now the largest phase
  of the CFB-only write, and emission is unchanged.
- **No claim about the same-length overlay path** (`render_copy_through`,
  record 0748's area) or about Rewrite.
- **Scope of the measurements:**
  - warm, in-memory, serial, on one host;
  - round one: four fixtures and the generated harness decks;
  - review round: seven fixtures, four generated many-mini-stream files, the
    harness selectors, and a corpus-wide instruction lane;
  - no cold-I/O, concurrency, RSS or cross-platform result;
  - peak live bytes are allocator ownership, not RSS.

## Follow-up candidates found, not changed

- **The reparse's per-stream map clearing (A5).** `validate_stream_allocations`
  clears a map sized to the whole FAT or MiniFAT for every stream
  (`SectorChainScratch::prepare_visited`). That is the one super-linear
  term left in the Reuse validation (×6.93 in Callgrind from 1,000 to 3,000
  mini streams), and it runs on every `OleFile::open` of the reader.
  `EndChainScratch`'s clearing, by chain or by table whichever is cheaper,
  would make it linear. It is left for its own record because it changes
  every reader's open path and needs the reader controls measured.
- **Planning.** `plan_reuse` is now about 45% of the CFB-only Reuse write
  (native samples). Most of it is:
  - the `place` closure's per-sector checks (12.8%);
  - per-element vector growth in `walk_chain`;
  - `SectorPool::high_water` scanning every sector in each fixed-point
    round;
  - path-keyed `BTreeMap` lookups.

  A planner defect still ends in a validation decline, so these are
  low-risk work for a following record.
- **The common editor renders into an empty `Vec`.** `render_with_layout`
  writes into `Cursor::new(Vec::new())`, which doubles its way up to the
  output size. The unattributed `memmove` under `write_to` (about 20% of the
  CFB-only write in both arms, with the probe's sink built the same way)
  includes that regrowth. Pre-sizing from the adopted source's length is a
  candidate.
- **Heap-layout sensitivity** now has a tested mechanism, glibc's trim and
  mmap thresholds. Later OLE2 lifecycle matrices should randomize layouts, or
  pin those tunables in an additional lane.

## Verification

`gates.txt` in the packet lists commands and exit codes for round one, and
`round2/gates.txt` for the review round.

**Review round, on `f2a57ac936`.** Every gate passed, run in the foreground
after measurement with `CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=6` and
`TMPDIR` under the scratch directory:

- `cargo fmt --all --check`.
- `cargo check --all-targets` for `litchi-cfb` and its eight dependents.
- Clippy `-D warnings` on the `litchi-cfb` library and all targets.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps`.
- Tests:

  | Crate | Passed | Ignored |
  |---|---:|---:|
  | `litchi-cfb` | 447 | 1 |
  | `litchi-ole-common` | 165 | 0 |
  | `litchi-doc` | 1,198 | 13 |
  | `litchi-ppt` | 1,229 | 11 |
  | `litchi-xls` | 1,474 | 1 |

  None failed; the ignored tests are pre-existing.
- Crate boundaries (64 packages, 241 declarations, 11 debt items),
  `non_iwork_gate verify`, and the structural perf-claims check (10 claims).

The facade and the remaining dependents' test suites were not rerun in the
review round; their check gate passed.

**Round one, on `776ad05175`.** All gates passed, run serially after measurement with
`CARGO_BUILD_JOBS=6` and `RUST_TEST_THREADS=6` under rustc 1.95.0:

- `cargo fmt --all --check`.
- `cargo check --all-targets` for `litchi-cfb` and its eight dependents:
  `litchi-ole-common`, `litchi-doc`, `litchi-ppt`, `litchi-xls`,
  `litchi-vba`, `litchi-ograph`, `litchi-crypto`, `litchi-sign`.
- `cargo check -p litchi --features doc,docx,ppt,pptx,xls,xlsx,xlsb,odt --all-targets`.
- Clippy `-D warnings` on the `litchi-cfb` library and all targets.
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-cfb --no-deps`.
- Tests:

  | Crate | Passed | Ignored |
  |---|---:|---:|
  | `litchi-cfb` | 440 | 1 |
  | `litchi-ole-common` | 165 | 0 |
  | `litchi-doc` | 1,198 | 13 |
  | `litchi-ppt` | 1,229 | 11 |
  | `litchi-xls` | 1,474 | 1 |
  | `litchi-vba` | 29 | 0 |
  | `litchi-ograph` | 67 | 0 |
  | `litchi-crypto` | 33 | 0 |
  | `litchi-sign` | 25 | 0 |
  | Facade (`litchi`, all OLE2/OOXML features) | 382 | 7 |

  None failed; the ignored tests are pre-existing.
- Crate boundaries: 64 packages, 241 declarations, 11 existing debt items.
- `non_iwork_gate verify`.
- Structural perf-claims check: 10 claims.

The harness is unchanged, so its tests and the coverage validator were not
rerun. No shared log, claim registry or coverage index is edited;
`log-sections.md` in the packet has the paste-ready sections.

## Cleanup

`cleanup.json` records both rounds.

- **Round one.** The coordinator committed the record and packet as
  `fc308af0d7` and then removed:
  - `targets/0749` (55G) and `targets/0749-before` (2.7G);
  - the `0749-before-src` checkout;
  - `scratch/0749`.
- **Review round.** After committing this record and packet, the
  implementer removes:
  - `targets/0749` (47G: the debug gate build, three probe builds and two
    harness builds);
  - the `0749-before-src` and `0749-v1-src` checkouts (9.5G each), with
    `git worktree remove --force`;
  - `scratch/0749` (640M).

The packet keeps sources, scripts, compressed raw reports and summaries
(15M). It holds no binaries, `perf.data`, Callgrind outputs or corpora; the
generated inputs regenerate byte-identically from the probe, and their
digests are recorded. The worktree and branch remain.

[Evidence packet and replay instructions](results/change-0749/README.md).
