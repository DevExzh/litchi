# 0594: one Deflate decoder per OOXML open, above a measured member threshold

Status: retained. `performance_claim: none` — this record carries exact,
deterministic allocation counts, `perf stat` isolation pairs, callgrind
instruction pairs, and paired timings reported beside their floors. Nothing here
is registered as a claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This implements item **ZIP-1** (rank 8) of change
[0587](0587-remaining-opportunity-survey.md), the mechanism that survey measured
at `crates/soapberry-zip/src/office.rs:4217-4222` and `1310-1318`.

**The first implementation of this record regressed one scenario and was held.**
The session was kept for every package; `docx_file_source_open` was then slower
at p50 in every paired direction. The regression is attributed below to minor
page faults — **+8 per fresh-process open, deterministic** — caused by the
retained workspace raising the open's heap high-water, and the change now keeps
a session only above a measured threshold. The attribution and the failed first
attempt are kept in this record because they are the reason the threshold exists.

One adverse row remains, on `xlsx_source_open`: −1.40% and −1.89% at p50 against
a 1.71% floor over 200 samples per leg. That selector's nine-member corpus is
below the threshold and **executes the pre-change path on both legs**, so the
residue is code layout rather than this mechanism; it is reported in *Paired
timing* and not explained away.

## What was changed

**`crates/soapberry-zip/src/office.rs`.** `IndexedArchive::read_with_accounting`
resolved a member name inline; that resolution moves into a private
`entry_id_for_name`, and `IndexedReadSession` gains `read` and
`read_with_accounting` by name, which call it. The session therefore admits a
name through exactly the code the one-shot read admits it through — the same
normalization, the same explicit-directory rejection, the same `FileNotFound`
identity — and then performs the read it already performed by entry ID.

**`crates/litchi-opc/src/pkgreader.rs`.** A crate-private
`SessionedArchive<'archive, R>` implements the private `ArchiveAccess` trait over
one `&IndexedArchive<R>` plus a session slot. Its `len`, `file_names`,
`metadata` and `read_stored_borrowed` are the plain positional archive's,
verbatim. It starts with **no** session. The trait gains one hook,
`note_relationship_member_count`, whose default body is empty;
`source_catalog` and `source_catalog_for_validation` call it with the count they
**already compute** for the `RelationshipParts` limit, before the first
structural read. `SessionedArchive` installs a session there if, and only if,
that count reaches `SESSION_MEMBER_THRESHOLD`. Below the threshold every read
takes `IndexedArchive::read`, the byte-identical path the package took before
this change, and the decision costs one integer comparison.

**`crates/litchi-opc/src/source_backed.rs` and `pkgreader.rs`.** The three places
that run the structural admission pass over a positional archive —
`SourceBackedPackage::from_read_at_*` (the ordinary open), `…_for_validation`
(the validation open) and `probe_package_catalog_from_reader_with_limits` (the
probe ingress) — hand `source_catalog` a `SessionedArchive` instead of the bare
archive. `source_catalog`, `walk_relationship_graph`, `load_rels_lazy` and
`read_structural_member` are otherwise untouched.

**`crates/litchi-opc/src/source_backed/batch.rs`.** `read_serial`, the serial
wave of `read_parts_ordered`, takes one session for the whole wave through
`SourceBackedPackage::sequential_read_session`, which returns `None` for a
managed package (change [0402](changes/0402-opc-overlay-decoder-reuse.md)'s rule)
and `None` for a wave shorter than the same threshold.

Nothing else changed. No public type, signature or default was added or altered;
`SessionedArchive` is `pub(crate)`, and the new `SourceBackedPackage` methods are
module-private.

### The threshold, and its rule

`SESSION_MEMBER_THRESHOLD = 16` relationship members.

> Keep the session only when it will avoid at least this many decoder
> constructions, so that the one workspace it retains is repaid at least
> sixteenfold in workspace it does not allocate and does not zero.

The margin is deliberately large rather than tuned to the crossover, because the
cost side is a heap-residency effect whose time cost is not predictable from the
package: the same +80,320 bytes of high-water cost **+8** minor page faults on a
20-member DOCX open, **+15** on a 132-member workbook, and **+1** on a
445-member PPTX. The benefit side is exactly countable — one construction per
member the session serves — so the rule is stated on the countable side and
demands a 16:1 margin over the bounded cost. Relationship members are counted
rather than archive members because they are exactly the members the session
serves: the pass reads `[Content_Types].xml` once and then one `.rels` per
relationship-bearing part.

The value is a scoped measurement, not a constant of nature. It is stated in the
source next to the numbers that produced it so that another allocator, platform
or corpus can re-derive it.

## Why it is sound

**The read is the same read.** `IndexedArchive::read_entry_with_accounting` is
already literally `self.read_session().read_entry_with_accounting(…)`. The only
difference a threaded session makes is which arm of one existing `match` supplies
the decoder: `decoder.reset(CountingReader::new(…))` instead of
`DeflateDecoder::new(CountingReader::new(…))`. `reset` on flate2 1.1.10 resets
the inflate state (`Decompress::reset`) and replaces the input reader, discarding
anything buffered. Store members never touch the decoder on either path. Below
the threshold no session exists and the pass executes the pre-change path.

**Error identity is unchanged, including after a refusal.** Name admission is now
one function for both paths. Payload refusals — bad CRC, corrupt Deflate bytes, a
truncated stream, a declared-size overrun or underrun — already had in-tree proof
that a session recovers
(`indexed_read_session_resets_after_crc_and_corrupt_deflate_failures`,
`archive_read_session_*`, `tests/deflate_reuse.rs`); this change adds the by-name
equivalent and a corpus-wide differential.

**Read-order independence holds (ADR 0005).** A session is not a cache: it
retains no payload, no metadata and no verdict, only decoder workspace. Every
value `source_catalog` computes is a function of member bytes, and the member
bytes are unchanged — checked member by member over the whole OOXML corpus, on
both sides of the threshold. The threshold itself is a function of the archive's
own member names, computed before the first read, so it cannot depend on read
order either.

**Limits and fences are untouched (ADR 0005).** No `ReadLimits` or
`ArchiveLimits` value, check or ordering moved. The hook is called after the
`RelationshipParts` check it reuses the count of, so a package that exceeds that
limit is refused exactly where it was refused before. On the batch path, cache
admission, the single-flight `Loader`/`Waiter`/`Bypass` arms, the
`source.ensure_current()` brackets and the budget reservations are the ones
`read_part_prepared` already used.

**Preservation is untouched (ADR 0006).** No output byte is produced or consumed
differently; nothing on the save, publication or precompressed-capture path was
modified.

**Managed budgets keep 0402's rule.** `sequential_read_session` returns `None`
whenever the package's cache carries an execution budget, so no decoder workspace
outlives a managed load.

**No new `unsafe`; nothing leaks.** `SessionedArchive` holds a `RefCell` and is
therefore `!Sync`. It is an operation-scoped stack value inside the three
admission functions, is never stored in a `SourceBackedPackage`, and never
crosses a thread; `SourceBackedPackage` gained no field. The batch session is
created inside the *serial* wave; the parallel waves keep the sessionless read.
The `RefCell` is reached through `try_borrow_mut` with a fall-back to the
archive's own one-shot read, so the "structural admission never re-enters a read"
reasoning is a performance assumption rather than a panic.

## Measured

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, rustc 1.95.0,
valgrind 3.26.0. Base `08d968f8ec7db27cf1187d01911fd08b9d014d91`. Both legs built
`--release --locked`, every measured process pinned to CPU 14.

### Where the extra cycles went: the attribution

The held implementation was profiled against the `docx_file_source_open` corpus
itself — 16,793,036 bytes, 20 archive members, **3 relationship members**,
SHA-256 `a4a2e492…`, extracted from the harness and retained — through the
selector's own `--filesystem-child`, which reports the timed region and its
`/proc` deltas, and through an isolation-pair probe that opens the same file 100
and 200 times.

| instrument | before | after (held) | finding |
| --- | ---: | ---: | --- |
| callgrind, Ir per open (isolation pair 20/40) | 1,381,264 | 1,279,346 | **−7.38%**, of which `memset` −94,178 |
| `perf stat`, instructions per open (100/200, `-r 5`) | 1,134,573 | 1,125,433 | −0.81% |
| `perf stat`, cycles per open | 293,930 | 294,851 | +0.31%, inside the counters' own 0.19–0.72% error |
| child `minor_faults`, 150 interleaved runs per leg | 92.4 mean, 92 p50 | 100.3 mean, 101 p50 | **+8 per open, at every quantile** |
| child `syscr` / `rchar` | 34 / 4,452 | 34 / 4,452 | identical |
| child `elapsed_ns`, same runs | 318,441 p50 | 321,872 p50 | **−1.08%**; min −0.44%, mean −1.22%, p95 −1.18% |

**The work fell and the cycles did not.** Callgrind's −7.38% is almost entirely
`memset` it charges per byte; the native instruction count fell only 0.81%,
because a hot 80 KiB zero-fill is a handful of AVX2 stores, and cycles were flat.
What changed deterministically was the process's heap high-water: the retained
80,320-byte workspace is memory the old code handed back between members, so the
catalog's own allocations had to take new pages. At the measured 3,431 ns / 8
faults ≈ 430 ns per fault, **the fault delta accounts for the whole 1.08%**.

Candidate causes named in the review and ruled out by the same instruments: the
`RefCell` borrow (one flag compare per read; total instructions and branches both
*fell*, −0.81% and −0.66%); `reset` being dearer than a fresh construction (it is
strictly less work — no `malloc`, no 80 KiB zero-fill — which is what the
callgrind `memset` line shows); inlining of the name-admission path (branch
misses fell 8.1%); and any I/O change (`syscr` and `rchar` are byte-identical).

Scale confirms the mechanism rather than contradicting it. The same instrument on
the `pptx_file_source_open` corpus — 445 members, **223 relationship members** —
showed **+1** minor fault and **+2.20%** faster at p50: the fixed cost is repaid
once the pass avoids enough constructions.

### The fix, and a first attempt that was worse

The first fix counted relationship members inside `SessionedArchive::new` by
scanning `file_names()`. That scan walks the index's heap-allocated member names
and, on a 20-member DOCX, cost **+4.83% cycles** and **+659% cache misses** per
open in the isolation pair — worse than the 1.08% it was meant to remove. It is
recorded here and not kept. The retained fix takes the count `source_catalog`
already computes for its `RelationshipParts` limit and passes it through the new
trait hook, which costs one comparison; the same isolation pair then reads
**+0.50% cycles, −0.16% instructions** — the pre-change path, within noise.

### Counts (deterministic, exact)

0587's probe, extended by a `read_parts_ordered` phase and a timing phase: an
in-memory `ReadAt` that logs every `read_at`/`version()` call, a counting global
allocator, `SourceReadPolicy::exact()`, `ReadLimits::default()`,
`SourceCacheLimits::default()`. One `DeflateDecoder::new` is **2 allocations /
80,320 bytes** on this host, measured directly on both legs.

Open, allocated bytes:

| fixture | rels members | structural reads | before → after | Δ | share |
| --- | ---: | ---: | --- | ---: | ---: |
| `ConditionalFormattingSamples.xlsx`, 132 members | 41 | 42 | 5,586,308 → 2,293,187 | −3,293,121 | **−58.95%** |
| `shapes.pptx`, 48 members | 20 | 21 | 2,199,830 → 593,429 | −1,606,401 | **−73.02%** |
| `comment.docx`, 10 members | 2 | 3 | 447,564 → 447,563 | — | **unchanged, below the threshold** |

The two above-threshold fixtures remove exactly *(structural reads − 1)*
constructions: 41 × 80,320 = 3,293,120 and 20 × 80,320 = 1,606,400. The DOCX
row's one-byte difference is the fixture path string, whose length differs
between the two checkouts; its allocation count is identical at 425.

**Positional requests, requested bytes and `version()` calls are identical on
every phase of every fixture** (open: 87 / 45 / 12 requests and 4 `version()`
calls; first cold part read: 2–3 requests and 5 `version()` calls).

Serial `read_parts_ordered` over every part of a freshly opened package (the row
covers that open plus the batch):

| fixture | parts | before → after | Δ | constructions removed |
| --- | ---: | --- | ---: | ---: |
| `…Samples.xlsx` | 90 | 11,249,114 → 2,092,634 | −9,156,480 | 114 (41 open + 73 batch) |
| `shapes.pptx` | 27 | 4,213,543 → 679,463 | −3,534,080 | 44 (20 + 24) |
| `comment.docx` | 7 | 1,009,991 → 1,009,991 | — | 0, below the threshold |

A **single** cold `PartView::data()` read is unchanged everywhere: 27 allocations
/ 86,291 bytes on the workbook, of which 80,320 is still one decoder.

### Instructions (callgrind, open plus one part, 132-member workbook)

| symbol | before | after | Δ |
| --- | ---: | ---: | ---: |
| PROGRAM TOTALS | 12,434,516 | 10,402,086 | **−2,032,430 (−16.35%)** |
| `__memset_avx2_unaligned_erms` | 2,234,582 (17.97%) | 304,261 (2.92%) | −1,930,321 (−86.4%) |
| `__memcpy_avx_unaligned_erms` | 989,772 | 962,393 | −27,379 |
| `_int_malloc` | 378,929 | 346,967 | −31,962 |
| `_int_free_merge_chunk` | 236,076 | 226,772 | −9,304 |

Callgrind charges `rep stosb` per byte, so the `memset` figure is an upper bound
on native cost — the attribution above is what that caveat looks like when it
bites.

### `perf stat` isolation pairs, per open

100 and 200 opens of the same package differenced and divided by 100, `-r 5`,
every counter at 100% coverage.

| event | DOCX corpus (3 rels, below) | PPTX corpus (223 rels, above) |
| --- | ---: | ---: |
| cycles | +0.50% | **−2.17%** |
| instructions | −0.16% | −1.53% |
| branches | −0.16% | −1.24% |
| branch misses | −2.06% | +0.39% |
| cache misses | −71.4% | −3.94% |

### Paired timing, with the floor beside it

**Interleaved, fresh process, the selector's own measured child.** The two legs
alternate sample by sample, so drift hits both equally; this is the instrument
that resolved the held regression.

| corpus | metric | before | after | change |
| --- | --- | ---: | ---: | ---: |
| DOCX, 3 rels (below) | `elapsed_ns` p50 | 316,252 | 308,512 | **+2.45%** |
| | min / mean / p95 | | | −1.02% / +2.13% / +2.18% |
| | `minor_faults` mean | 92.4 | 93.2 | +0.9 (was +7.9 when held) |
| PPTX, 223 rels (above) | `elapsed_ns` p50 | 2,959,376 | 2,896,805 | **+2.11%** |
| | min / mean / p95 | | | +2.62% / +3.93% / +12.49% |
| | `minor_faults` mean | 511.9 | 509.9 | −2.0 |

`syscr` is identical on both corpora (34 and 675 reads).

**In-process, through the retained probe**, 10 warmups then 40 timed opens per
leg, order A1 B1 B2 A2, CPU 14:

| fixture | stat | A1→B1 | A2→B2 | A/A floor |
| --- | --- | ---: | ---: | ---: |
| `…Samples.xlsx`, 41 rels (above) | p50 | **+4.88%** | **+5.34%** | 0.14% |
| | mean / p95 / p99 / min | +5.33% / +6.30% / +5.93% / +5.89% | +5.72% / +5.84% / +2.43% / +6.92% | ≤0.22% |
| `comment.docx`, 2 rels (below) | p50 | +2.30% | +1.90% | 0.41% |
| | mean / p95 / min | +3.05% / +2.06% / +2.22% | +1.65% / +2.43% / +2.48% | ≤1.58% |

The `comment.docx` row executes the pre-change path on both legs; its ~2% is code
layout, not the mechanism, and is reported rather than claimed.

**Harness ABBA**, order A1 B1 B2 A2, each leg a separate process pinned to CPU
14, the harness's own fresh-child isolation for the filesystem selectors, p50
unless stated:

| selector / cache | samples | A1→B1 | A2→B2 | A/A floor |
| --- | ---: | ---: | ---: | ---: |
| `docx_file_source_open` warm | 40 | +4.98% | **−0.74%** | 5.52% |
| `docx_file_source_open` cold-requested | 40 | +5.17% | **−0.20%** | 7.61% |
| `pptx_file_source_open` warm | 40 | **+4.19%** | **+2.76%** | 3.22% |
| `pptx_file_source_open` cold-requested | 40 | **+5.50%** | **+3.58%** | 3.12% |
| `xlsx_source_open` (medium) | 40 | −3.48% | −0.48% | 1.81% |
| `xlsx_source_open` (medium), focused repeat | 200 | **−1.40%** | **−1.89%** | 1.71% |

`docx_file_source_open` is the scenario that held this change. Its adverse
direction is now −0.20% and −0.74% against floors of 5.52% and 7.61% — inside the
floor, and two orders of magnitude smaller than the held implementation's −4.12%
to −11.56%. `pptx_file_source_open`, the only harness corpus above the threshold,
is favourable in all four directions.

**One adverse row remains and is reported rather than pooled.**
`xlsx_source_open` uses a nine-member synthetic workbook with two relationship
members, so it is below the threshold and **executes the pre-change path on both
legs**; the focused 200-sample repeat puts it at −1.40% and −1.89% at p50 against
a 1.71% floor, about 0.7 µs of a 44 µs operation. It cannot be this change's
mechanism, because the mechanism does not run; the residue is code layout in a
binary that carries added, unexecuted code. No further attribution was taken and
none is claimed.



## Correctness evidence

**A corpus-wide differential on the changed seam.**
`sessioned_structural_reads_match_one_shot_reads_across_the_ooxml_corpus`
(`crates/litchi-opc/src/pkgreader.rs`) walks `test-data/ooxml`, indexes every
OOXML package it finds — **179 packages, 4,215 members** — drives the same hook
`source_catalog` drives with the same count, and for each member compares the
session-bearing `ArchiveAccess` against the archive's own one-shot read: `len`,
`file_names`, `metadata`, `read_stored_borrowed` and the payload, asserting byte
equality on success and identical error rendering on failure. One session is
reused across every member of a package. The test asserts that the corpus
exercises **both** sides of the threshold, which it does: **15 packages keep a
session, 164 do not**.

**The threshold rule has its own test.**
`a_session_is_kept_only_for_packages_above_the_member_threshold` builds archives
carrying 0, 1, 15, 16 and 20 relationship members beside one
`[Content_Types].xml`, asserts that no session exists before the pass reports its
count, that the content-types member is not counted, that a session is installed
at exactly and above the threshold and not below it, and that the payload bytes
and the `FileNotFound` identity are the archive's own on both sides.

**A by-name refusal-recovery test in the ZIP crate.**
`indexed_read_session_reads_by_name_exactly_as_the_one_shot_read`
(`crates/soapberry-zip/src/office.rs`) builds a four-member archive — one Store,
one Deflate member with a corrupted central CRC, two ordinary Deflate members —
reads every member by name through one session comparing payload bytes,
`ZipOperationAccounting` and error rendering against the one-shot read, re-reads
in a different order, re-triggers the refusal, and checks `absent.bin`,
`deflated.bin/` and `stored.bin\` against the archive's own `FileNotFound`.

**Gates** (tails in `results/change-0594/gates.txt`):

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p soapberry-zip --all-targets` | clean (workspace lints deny) |
| `cargo clippy -p litchi-opc --all-targets` | clean |
| `cargo test -p soapberry-zip` | 588 passed, 0 failed |
| `cargo test -p litchi-opc` | 666 passed, 0 failed |
| `cargo doc -p soapberry-zip --no-deps` | clean (rustdoc lints deny) |
| `cargo doc -p litchi-opc --no-deps` | clean |

The pre-existing suites that cover this seam pass unchanged, including
`litchi-opc`'s `source_backed_batch_reuse`, `source_backed_topology`,
`source_backed_reader`, `source_read_ahead`, `source_part_splice*`, `validation`
and `tolerated_archive_junk`, `soapberry-zip`'s `deflate_reuse` and
`descriptor_read_once`, and the crate's scoped-thread concurrency tests.

## Validation preserved

The validation open takes the same hook and the same threshold, and its phase
provenance (`ValidationCatalogPhase::{Ingress, Catalog, LoadedRelationships}`) is
unchanged: the wrapper is passed where the archive was and every `map_err` stays
where it was. Validation still performs no mutation and still reads exactly the
members it read before. No typed refusal moved: the refusals reachable from this
seam are `FileNotFound` (unchanged admission code, now shared), the ZIP payload
refusals (unchanged reader, only the decoder's provenance differs) and the OPC
limit refusals (untouched, and the hook is called after the one whose count it
borrows).

## Limitations

* **Reach.** With the threshold, the change applies to packages carrying at least
  16 relationship members: **15 of the 179 OOXML fixtures in this repository
  (8.4%)**, whose relationship-member counts run 2 (median 3, p90 15) to 41.
  Those are the multi-slide presentations and multi-sheet workbooks — the large,
  slow opens where the 58.95% and 73.02% reductions were measured — and the
  remaining 164 fixtures are byte-for-byte on their previous path. No claim is
  made for them, in either direction.
* **The threshold is a scoped measurement.** 16 is a 16:1 margin over a bounded
  cost, taken on one host with one allocator. A different allocator or platform
  would need it re-derived; the source states the rule so that it can be.
* **No pooled session for a single cold part read.** A cold `PartView::data()`
  read still builds one decoder: 80,320 bytes, measured and unchanged. Retaining
  a session on the package is not expressible safely —
  `IndexedReadSession<'archive, R>` holds `&'archive IndexedArchive<R>` and a
  decoder over `ZipReader<&'archive R>`, while `SourceBackedPackage` *owns* the
  archive, so a pooled session would be a self-referential borrow. Pooling only
  the inflate engine is not available either: flate2 1.1.10 exposes no
  constructor that installs a retained `Decompress` into `read::DeflateDecoder`.
* **Peak transient above the threshold.** A package that keeps a session holds
  80,320 bytes across the pass; measured as +15 minor page faults per
  fresh-process open of the 132-member workbook and −2 on the 445-member PPTX. No
  peak-RSS or high-water claim is made.
* **Instruction counts rank work, not latency**, and callgrind's `memset` figure
  is a per-byte upper bound — this record shows exactly how far apart those can
  be.
* **Not claimed:** any saving on the eager `LazyArchiveReader` ingress, on the
  parallel batch waves, on managed packages, or on the verified/streaming
  `stream_to` path. Nothing is claimed for cold caches, physical devices, range
  sources, remote transports, concurrency scaling, RSS or other platforms.

## Retained evidence

[`results/change-0594/README.md`](results/change-0594/README.md) — the two probe
sources, the before and after count outputs, the callgrind tables, the `perf
stat` isolation pairs, the interleaved child measurements for both corpora and
both implementations, the ABBA timing reports with their drivers, the held
first implementation as a patch, `gates.txt`, `decision.json` and
`log-sections.md`.
