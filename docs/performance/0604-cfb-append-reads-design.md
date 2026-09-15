# 0604: the OLE2 whole-stream zero-fill is 0.6-2.9% of an open in native cycles, not 9-32%; the appending design is frozen and deliberately not implemented

Status: retained, design only. `performance_claim: none` — the numbers below are
paired-median cycle and instruction counts reported as evidence, not registered
as claims. **No file under `crates/` was modified by this change.**

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was changed

Nothing in production. This record carries three things:

1. the **measured native ceiling** of the whole-stream zero-fill on the eager
   DOC open, the eager and source-backed PPT opens and the source-backed XLS
   open, each on an owned in-memory source, against a micro-benchmark of the
   same byte counts on this host;
2. the **frozen design** of the appending read shape — a provided `ReadAt`
   method, a single chain walk with two sinks, tail-only zeroing in the sector
   helpers, `try_reserve` before every append, and the parity gate — so that a
   later batch has the shape without re-deriving it;
3. the **recommendation not to implement it**, because the ceiling is inside the
   floor on every fixture measured, which is the falsification criterion change
   0587 wrote for this item (CFB-1, rank 22) when it ranked it.

## The ceiling, measured first

Change 0587 sized this item from retained callgrind counts: `memset`
instructions per open equal the bytes slurped, 1,595,476 Ir against 1,595,422 B
on the 1.6 MB DOC, and the two slurping callers are 9-32% of the DOC and PPT
opens. That share is a `rep stosb` upper bound: callgrind prices the instruction
once per byte, where the hardware retires it once per run. This record prices it
natively.

### Bytes actually zero-filled per open (measured, exact)

Counted by instrumenting every `u8` zero-fill site in `litchi-cfb` and the XLS
globals buffer and running one open per fixture on an owned in-memory source
(`results/change-0604/zero-fill-byte-counts.txt`, patch retained).

| scenario | fixture (bytes) | zero-filled per open | of which whole-stream slurp | sector-helper double-zero |
| --- | ---: | ---: | ---: | ---: |
| DOC eager open | `ca.kwsymphony…doc` (1,619,457) | 1,623,070 | 1,595,422 in 3 streams | 26,624 |
| DOC eager open | `FloatingPictures.doc` (335,360) | 323,850 | 315,658 in 3 streams | 6,144 |
| PPT eager open | `45543.ppt` (385,024) | 360,903 | 353,735 in 3 streams | 6,144 |
| PPT source-backed open | `45543.ppt` (385,024) | 322,788 | 315,620 in 2 streams | 6,144 |
| XLS source-backed open | `ConditionalFormattingSamples.xls` (1,402,368) | 565,713 | 0 (no CFB slurp) + 551,377 XLS globals in 20 fills | 12,800 |

The XLS row confirms the survey's reading: a source-backed XLS open does not
slurp a stream through `litchi-cfb` at all; its zero-fill is
`GlobalsBuffer::ensure` in `litchi-xls` (XLS-7), and the CFB contribution is
14,336 bytes.

### What the zero-fill costs, natively

`perf stat` cycles and instructions, isolation pairs at 10 and 110 samples,
median of 11 repetitions each, `taskset -c 19`, release build, owned in-memory
sources (`results/change-0604/ceiling-base.txt`).

Two models of the term are reported, because they bound it from different sides:

- **`memset` alone** — `try_reserve_exact` + `resize(n, 0)`, exactly
  `try_zeroed_vec`/`try_filled_vec`. This is what the brief asked for.
- **`zero+copy` minus `append`** — `reserve; resize(n,0); copy_from_slice`
  against `reserve; extend_from_slice`. This is what the design would actually
  change, and on the 1.6 MB fixture it is the larger of the two, because the
  zero pass and the copy pass each write the whole buffer out to memory where
  the appending shape writes it once.

The ceiling per scenario is the larger of the two.

| scenario | open cycles/op | open Ir/op | `memset` cycles | `zero+copy − append` cycles | **ceiling, cycles** | **ceiling, Ir** |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DOC eager, docbig | 2,310,748 | 5,758,552 | 31,315 | 43,399 | **1.88%** | 0.78% |
| DOC eager, docmid | 915,516 | 3,421,321 | 5,194 | 3,842 | **0.57%** | 0.29% |
| PPT eager, pptmid | 211,456 | 573,970 | 6,010 | 5,054 | **2.84%** | 1.88% |
| PPT source-backed, pptmid | 197,940 | 554,294 | 5,303 | 4,751 | **2.68%** | 1.69% |
| XLS source-backed, flagship | 317,682 | 1,134,885 | 9,339 | 9,307 | **2.94%** | 1.39% |

The allocation itself is free at this scale: `try_reserve_exact` without the
`resize` measured −509 to +1,014 cycles per operation, i.e. inside the noise, on
every byte count. The zero-fill is a store pass, not an allocation.

### The A/A floor, in the same window

Two floors were taken while the ceiling runs were interleaved with them:

- on the ceiling metric itself (isolation-pair cycles), the DOC open measured
  2,319,335 against 2,309,157 (0.44% apart) and the source-backed PPT open
  200,101 against 200,688 (0.29% apart);
- on wall-clock paired timing, 60 samples per leg in `A1 B1 B2 A2` order, each
  sample the probe's mean over 30 iterations after 3 warm-ups, pinned to CPU 19
  (`results/change-0604/aa-floor.txt`):

| scenario | leg A p50 / p95 / p99 (ns) | leg B p50 / p95 / p99 (ns) | p50 B/A − 1 |
| --- | --- | --- | ---: |
| DOC eager open | 577,366 / 580,814 / 582,092 | 576,997 / 581,961 / 583,906 | −0.06% |
| PPT source-backed open | 57,792 / 59,142 / 59,468 | 57,795 / 59,370 / 59,501 | +0.00% |
| XLS source-backed open | 88,704 / 160,709 / 161,700 | 88,939 / 193,419 / 194,506 | +0.26% |

The floor did not exceed 5% at p50 on any scenario, so the timings are usable.
The XLS p95 and p99 are 1.8× to 2.2× its p50 on **both** legs: seven other
agents were building on this host throughout, and that tail is host noise, not a
property of the scenario.

### The decision

Change 0587 wrote the falsification criterion for CFB-1 itself: "falsified if a
cycles A/B of the DOC and PPT opens on an owned source sits inside the 4% p50
floor". Measured, the ceiling is 0.57% to 2.94% of the open, inside 4% on every
fixture, and the realizable saving is strictly below the ceiling because the
appending shape adds per-run bookkeeping to a walk that today writes one
contiguous buffer. **The design is frozen and not implemented.**

Two further facts push the same way and are the reason the result is not
"measure again with a tighter harness":

- **`FileSource` gets nothing.** No stable positional API accepts uninitialized
  memory — `std::os::unix::fs::FileExt::read_at` takes `&mut [u8]` and
  `BorrowedBuf` is unstable — so the default provided method keeps today's
  zero-fill for every file-backed source. Every facade route that opens a `.doc`,
  `.ppt` or `.xls` from a path goes through `FileSource` or a `File` reader. The
  ceiling above is measured on the owned in-memory shape, which is the harness.
- **The callgrind share this item was ranked on overstates the native cost by
  more than an order of magnitude.** The retained figure for the DOC open is
  1,595,476 `memset` Ir, 32.15% of the open; natively the same work retires
  44,925 instructions, 0.78% of the open, and costs 1.88% of its cycles. That is
  a 35× instruction-count overstatement, and it is a general caution for every
  bulk-copy share in this record set, not a fact about this item alone.

## The frozen design

Everything below is the shape a later batch should implement if a scenario is
ever found where the ceiling clears the floor — a source far larger than this
corpus's 1.6 MB maximum, or an owned-source path that slurps repeatedly. It is
written to be implemented as stated, not as a sketch.

### Part 1 — `litchi-core`: one provided `ReadAt` method

```rust
/// Appends exactly `len` bytes read at `offset` to `output`.
///
/// The default reserves, zero-extends and fills, so every existing
/// implementation keeps today's behaviour unchanged. An adapter that owns
/// contiguous bytes overrides it and never writes a zero.
fn read_exact_at_appending(
    &self,
    offset: u64,
    len: usize,
    output: &mut Vec<u8>,
) -> io::Result<()> {
    let start = output.len();
    output
        .try_reserve(len)
        .map_err(|_error| io::Error::from(io::ErrorKind::OutOfMemory))?;
    output.resize(start.saturating_add(len), 0);
    match self.read_exact_at(offset, &mut output[start..]) {
        Ok(()) => Ok(()),
        Err(error) => {
            output.truncate(start);
            Err(error)
        },
    }
}
```

`OwnedSource`, `SliceSource` and `litchi-cfb`'s `OwnedArcSource` override it
with a bounds-checked `extend_from_slice` of the borrowed range, raising the
identical `UnexpectedEof` error — `io::Error::new(ErrorKind::UnexpectedEof,
"positional source ended before the requested range")` — that `read_exact_at`
produces today when `read_slice_at` returns a short count.

Three properties the design fixes:

- **Appending is atomic.** On any error the output is truncated back to its
  entry length, so a failed append leaves no partial bytes for the chain walk to
  reason about.
- **The default is behaviour-identical**, so this is an additive change to a
  public trait and no out-of-tree implementor is broken. It adds no dependency
  and exposes no archive type, lock or executor: `docs/GOAL.md` rule 11 is
  satisfied because the method moves bytes, not grammar.
- **No `unsafe`.** Rule 10 is untouched; the whole point of the shape is to make
  the zero avoidable without `MaybeUninit`.

### Part 2 — `litchi-cfb` shared reads: one chain walk, two sinks

`read_chain_into` (`shared.rs:2538-2611`) validates while it walks, and the
messages and their order are the contract: `invalid {table} stream chain`,
`{table} chain ends before its declared length`, `{table} chain exceeds its
declared length`, `Sector {n} is outside the file`, plus the CFB overflow
refusals. The walk must therefore stay single-sourced. Parameterize its output:

```rust
enum ChainSink<'a> {
    Slice { output: &'a mut [u8] },
    Append { output: &'a mut Vec<u8> },
}
```

with one operation, "take this run": `Slice` keeps today's
`read_sector_run(run_start, &mut output[start..end])` verbatim; `Append`
computes `position` exactly as `read_sector_run` (`shared.rs:2613`) does, refuses `position >=
file_size` with the identical `Sector {n} is outside the file` message, computes
`present = min(file_size − position, want)`, calls
`source.read_exact_at_appending(position, present, output)`, and then
zero-extends by `want − present`. **That zero-extension is the truncated-final-
sector semantics**, and it is the only zero the shape writes.

`read_fat_stream` becomes `try_reserve_exact(size)` on a fresh `Vec` followed by
the walk with `ChainSink::Append`. The up-front exact reserve is not an
optimization: it keeps the allocation refusal exactly where it is today — one
typed `OleError::allocation("FAT stream data")` raised before any read — instead
of letting it surface as an `io::ErrorKind::OutOfMemory` from inside
`litchi-core` partway through the chain. `read_stream_range_hinted`, the hinted
cursor (0579, 0585) and `read_minifat_stream` keep `ChainSink::Slice` and do not
change.

Rule 12 is satisfied twice over: the up-front `try_reserve_exact` bounds the
allocation before any source byte is read, and the per-run `try_reserve` inside
the trait method is a no-op against that capacity but remains as the defence for
any other caller.

### Part 3 — `litchi-cfb` eager reads, the DOC site

`OleFile<R: Read + Seek>` (`file.rs:222`) has no `ReadAt`, so Part 1 does not
reach it — and it is the **largest** payer: 1,595,422 of the DOC open's
1,623,070 zero-filled bytes are `read_stream_from_fat` (`file.rs:2115`, the
zero-fill at `:2133`). Two shapes exist; the design freezes the first and
records the second as listed, not proposed:

- **(3a)** replace `try_filled_vec(size, 0)` + `read_exact(&mut buf[..present])`
  with a reserved `Vec` and `(&mut self.reader).take(present).read_to_end(&mut
  data)`. `std`'s own `Read` implementations for `Cursor<_>`, `&[u8]` and
  `File` provide `read_buf`, which is every reader a production caller supplies,
  so no zero is written; a reader without `read_buf` degrades to exactly today's
  cost. That specialization is a `std` implementation detail and must be
  re-verified against the toolchain in use before the part is relied on. Two
  obligations: `read_to_end` reports a short read as `Ok(n)` where `read_exact`
  raises `UnexpectedEof`, so the helper must compare `n` with `present` and
  raise the identical `io::Error`; and the reserved capacity must be proven
  sufficient so `read_to_end` never grows the buffer past it. This is the part
  of the design with the weakest error-identity proof and it needs a differential
  over the malformed corpus before it lands.
- **(3b)** give `ReadAtCursor` (`shared.rs:3151`) a `Read::read_buf`
  implementation. It helps only `SharedOleFile::open_with_limits`, which reads
  the header, FAT, directory and MiniFAT and slurps no stream, so it is listed
  and not proposed.

### Part 4 — tail-only zeroing in the sector helpers (CFB-4, never lands alone)

`read_sector_into` (`file.rs:1633-1660`, `buffer.fill(0)` at `:1647`) and
`read_sector_run_into` (`file.rs:1700-1760`, `run.fill(0)` at `:1744`) fill the
whole buffer and then overwrite `[..present]`;
`sector_batch_scratch` (`file.rs:1684`) is zero-allocated on top, so each run is
zeroed twice. Reading first and zeroing `[present..]` afterwards — as
`file.rs:1792` already does — preserves the success path exactly. On the
**failure** path it does not: the buffer would keep the previous sector's bytes
where today it holds zeros. Every current caller returns the error without
reading the buffer, and the design requires that to be re-proven caller by
caller before it lands. Measured size: 26,624 bytes of the DOC open's 1,623,070,
6,144 of the PPT open's 360,903, 12,800 of the XLS open's 565,713 — 1.6% to 2.3%
of a term that is itself inside the floor. Change 0587 is right that it must
never land alone.

### Part 5 — the XLS globals buffer (XLS-7)

`GlobalsBuffer::ensure` (`workbook/source.rs:1608-1660`, the `resize` at
`:1645`) grows `self.bytes` with
`try_reserve_exact(finish − start)` + `resize(finish, 0)` and then fills
`[start..finish]` through `read_stream_range_hinted`. The appending form needs a
`SharedOleFile::read_stream_range_appending`, which is Part 2's sink applied to
the hinted walk, and it must preserve 0579's resumable chain hint and 0585's
cursor resume semantics unchanged. Measured: 551,377 bytes zeroed across 20
fills on the flagship, 2.94% of that open in cycles.

Adjacent observation, **not sized and not claimed**: `try_reserve_exact` asks
for exactly `finish − start` more bytes, so each of the 20 fills reserves to an
exact capacity rather than growing geometrically. Whether that costs a copy per
fill depends on whether the allocator can `mremap` the block. This record did
not measure it and makes no claim about it; it is recorded so the next reader of
this site does not have to rediscover it.

### The parity gate the design must pass

1. the three truncated-final-sector tests —
   `fat_stream_read_zero_fills_a_truncated_final_sector`,
   `batched_fat_reads_zero_fill_a_truncated_final_sector` and
   `direct_open_zero_fills_a_truncated_final_root_sector_like_the_cache`
   (`file.rs:3875`, `file.rs:4369`, `shared.rs:5292`) — which exist and pass at
   this base (`results/change-0604/gates.txt`) — plus
   `read_sector_into_zero_fills_a_truncated_final_sector` (`file.rs:4772`) for
   Part 4;
2. a differential over every OLE2 fixture under `test-data/`: every stream's
   bytes compared for byte identity between the old and new readers;
3. the malformed corpus compared by `OleError` `Display` string, including a run
   starting at or past end of file, which is precisely the regression change
   0570 declined a helper to avoid — "it would have silently zero-filled where
   the old path raised a typed error";
4. `cargo fmt`, `clippy`, `cargo test` and `cargo doc` on `litchi-core`,
   `litchi-cfb` and `litchi-xls`.

### What the design does not touch

`read_stream_range_hinted`'s existing slice path; the MiniFAT copies
(`shared.rs:2146`, `:2480`) and `load_ministream`'s `Vec → Arc<[u8]>` (CFB-5);
`clone_minifat_waiter_payload`; the directory buffer (`file.rs:1025`);
`sector_roles`; `collect_sector_chain`'s per-read bitset (CFB-3); the writer and
the overlay; every ADR 0005 mandatory validation; and `FileSource`, which keeps
the default and therefore today's behaviour.

## Why it is sound

**Invariants.** The zero-fill exists for exactly one reason: a stream whose
declared size runs past the end of the file keeps zeroes in its truncated final
sector. The design preserves that by zero-extending `want − present` bytes at
the one place the truncation is known, rather than pre-zeroing the whole buffer
to cover it. No other byte of any stream is observably zero today, because
`read_chain_into` overwrites `[..present]` of every run it walks.

**Error identity.** Three refusals are load-bearing and all three are held by
construction: `Sector {n} is outside the file` for a run starting at or past end
of file (0570); the chain-length pair, `ends before` and `exceeds`, which the
single shared walk continues to raise in the same order; and
`OleError::Allocation` naming `"FAT stream data"`, which the up-front exact
reserve keeps ahead of any read. The appending trait method is atomic on error,
so no refusal can leave partially appended bytes behind.

**ADR reading.** ADR 0005's mandatory structural validations are untouched:
nothing here makes `validate_stream_allocations`,
`validate_physical_sector_layout` or `collect_exact` closure-proportional, which
is what 0574 opportunity 6 rejected. ADR 0006's preservation and ownership
boundaries are untouched: no read crosses into a sector another stream owns,
because the walk is unchanged; no output byte changes; no `SourceVersion` fence
moves. ADR 0003 is not engaged, because nothing is published.

**`docs/GOAL.md` rules.** Rule 10 — no new `unsafe`, which is the whole reason
for the trait shape. Rule 11 — a `litchi-core` trait addition that moves bytes,
adds no dependency and leaks no archive type, lock or executor. Rule 12 —
`try_reserve` before every append, the typed allocation refusal kept where it is,
and no limit, budget, cancellation point or malformed-input defence weakened.

**Which contracts are untouched.** The public `ReadAt` gains a provided method,
so no implementor must change. `OleFile`, `SharedOleFile`, `DirectoryEntry` and
`SourceBackedWorkbook` keep their signatures. No output byte of any writer or
save path is affected, because no writer path is in scope.

## Measured

Everything in "The ceiling, measured first" above is this section's content and
is not repeated. In tiers:

- **measured** — the zero-filled byte counts per open, exact, from instrumented
  single opens; the per-operation cycles and instructions of the five opens; the
  cycles and instructions of the `memset`, `zero+copy`, `append` and
  allocation-only micro-benchmarks at each fixture's byte counts; both A/A
  floors.
- **modelled** — the ceiling percentages, which divide a micro-benchmark of the
  same byte counts by a measured open. The micro-benchmark reproduces the
  allocation shape (`try_reserve_exact` then `resize`/`extend_from_slice`, a
  fresh buffer per iteration) but not the interleaving of the zero-fill with the
  chain walk, so it measures the term in isolation rather than in place.
- **unknown** — the realizable saving. It is bounded above by the ceiling and is
  not measured, because the change was not implemented.

Scope of every figure: scenario as named; corpus the five fixtures named;
machine AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws, CPU 19, with
seven other agents building concurrently; build `--release` with `debug = 1`,
rustc 1.95.0; metric `perf stat` `cycles` and `instructions`, and wall-clock
nanoseconds for the A/A.

## Correctness evidence

No production code changed, so there is nothing to test that was not already
tested. What was run, at the base commit on the clean worktree, is the parity
gate the design names plus the crate gates for the three crates the design would
touch:

| gate | result |
| --- | --- |
| `cargo fmt --all --check` | clean |
| `cargo clippy -p litchi-cfb -p litchi-core -p litchi-xls --all-targets --locked` | clean (workspace lints are `deny`) |
| `cargo test -p litchi-cfb --release --locked` | 322 + 13 + 6 passed, 12 passed 1 ignored doctests, 0 failed |
| the three truncated-final-sector parity tests, named | 3 passed, 0 failed |
| `cargo doc -p litchi-cfb -p litchi-core -p litchi-xls --no-deps --locked` | clean (rustdoc lints are `deny`) |

The instrumentation used for the byte counts was applied, captured and reverted;
the working tree was verified clean before the gates were run and before the
commit. The patch is retained so the counts are reproducible.

## Validation preserved

Nothing moved. Every ADR 0005 mandatory validation runs where it ran;
`validate_stream_allocations` and `validate_physical_sector_layout` still walk
every stream's chain and every physical sector at open; every typed limit,
budget and cancellation point is where it was; no refusal moved earlier or later;
no output byte changed. This record adds a design and five measurements.

## Limitations

- **The ceiling is a ceiling.** It is what the zero-fill costs in isolation, not
  what removing it would return. The appending shape adds per-run bookkeeping
  and an extra bounds check per run to a walk that today writes into one
  contiguous slice; the realizable saving is strictly smaller and is not
  measured.
- **Owned in-memory sources only.** Every open measured here wraps bytes already
  in memory. That is the shape the design can help and the shape the harness
  uses; it is not the shape `Package::open(path)`,
  `SourceBackedWorkbook::from_path` or `SourceBackedPackage::from_path` use, and
  on those the design is a no-op by construction. No file-backed leg was
  measured, because there is nothing for the design to change there.
- **Five fixtures, one host, one build.** The corpus maximum is 1.6 MB with no
  DIFAT sector and no 4,096-byte sector (0587's census over 212 fixtures). A
  synthetic large or v4 fixture could move the ceiling; none exists, and this
  record did not build one.
- **The micro-benchmark is not the open.** It allocates and fills in a tight
  loop, so the allocator reuses the same warm heap block after the first
  iterations, where a single cold open would pay page faults on top. That
  affects both of its legs equally and both directions of the comparison, but it
  means the absolute cycle figures for the `slurp-*` cases are a warm-heap
  number, not a cold-start one.
- **Not claimed:** any speedup, any regression, any file-backed, range-source,
  cold-cache, peak-RSS, allocation-count or cross-platform result; any statement
  about DOC, PPT or XLS opens outside the five fixtures named; any size for the
  `try_reserve_exact` realloc observation in Part 5; any conclusion about
  CFB-2 through CFB-6, XLS-6 or the other items 0587 ranks near this one.
- **What is left open.** The design is frozen but unexercised: no differential
  harness over the OLE2 corpus was written, because nothing was implemented to
  differentiate. Part 3a's error-identity proof is the piece a future batch
  should do first, because it is the only part that changes how a short read is
  reported.

## Retained evidence

[`results/change-0604/README.md`](results/change-0604/README.md) — the probe
source, the instrumentation patch, the capture scripts, the raw `perf stat`
outputs for every case, both A/A floors, the byte counts, the gate tails,
`decision.json` and `log-sections.md`.
