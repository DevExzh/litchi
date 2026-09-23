# 0752 — a one-pass CRC-32 stage, handle-free budget reservations and run-wise escaping cut streaming DOCX creation by 47% with every output byte unchanged

Status: retained, implemented in `litchi-core`, `soapberry-zip`, `litchi-docx`
and `litchi-xlsx`. `performance_claim: none` — the paired timings, instruction
counts and allocation counts below are evidence, not a registered claim.

OLE2 and OOXML remain the active priority; ODF stays deferred and iWork is
excluded. Base `6d989cad63`; branch `perf/0752-streaming-writer-small-write-batching`,
commits `9ce59dc565` (budget), `709efbafab` (ZIP CRC stage) and `1167398a3e`
(DOCX and XLSX writers). The coordinator's task: make the streaming OOXML
creation writers stop paying per-tiny-write costs in budget accounting and
CRC-32, keeping every budget limit, typed refusal and output byte identical.

**Result.** `docx_streaming_create` falls from 91.474 ms to 48.596 ms on the
large corpus (−46.66%, 95% CI [−47.25%, −45.24%]) and from 5.754 ms to
3.094 ms on the medium one (−46.30%), executing 27.5% fewer instructions per
iteration. `xlsx_streaming_create` improves 1.94% (large) and 3.34% (medium);
`pptx_streaming_create` 1.45% (large, CI spanning zero) and 1.86% (medium).
Every published digest is identical in the 112 measured processes that report
one, and a differential probe built against both trees prints identical
transcripts over 28,266 limit, cancellation and sink-failure scenarios. The
task's second direction — coalescing small writes before the compressor — is
**not byte-transparent** for the workspace's zlib-rs, measured below, so only
the CRC-32 is staged.

## What was changed

* `crates/litchi-core/src/budget.rs`
  * One private `charge_chain(leaf, resource, amount)` now does the work of
    `Budget::consume` and of the old private `Budget::charge`: one atomic
    check-and-add per hierarchy level, innermost first, walking the immutable
    parent chain by reference, releasing the levels already charged when a
    level refuses. The refusal (`ResourceLimit` with the refusing level's
    scope, the observed value and its limit) is built exactly as before.
  * `Reservation` holds one `Arc<Node>`, the node it charged, instead of a
    `SmallVec<[Arc<Node>; 4]>` with one clone per level. The old `charge`
    cloned the handle two or three times per level and dropped it once per
    level; `commit` and `Drop` then dropped one handle per level. Now `reserve`
    takes one handle, and `commit` and `Drop` release amounts by walking the
    chain by reference (`release_chain`). A zero remainder is no longer
    "released" by an atomic update that changed nothing.
  * New `Budget::reserve_scoped(&self, resource, amount) ->
    Result<ScopedReservation<'_>, ResourceLimit>` and `ScopedReservation`
    (`amount`, `resource`, `commit`, `Drop`): the same charge, refusal, commit
    and release as `reserve`/`Reservation`, but the token borrows the budget,
    so making and dropping it touches no reference count.
* `crates/litchi-core/src/execution.rs`: `ExecutionContext::reserve_scoped`,
  which checks cancellation first exactly as `ExecutionContext::reserve` does.
  `crates/litchi-core/src/lib.rs` re-exports `ScopedReservation`.
* `crates/soapberry-zip/src/writer.rs`
  * A private `CrcStage` in `ZipDataWriter`. When enabled, each accepted write
    of at most 1,024 bytes is copied into a 4 KiB stage, which is folded into
    the CRC-32 in one pass when the next write would overflow it, before a
    longer write (checksummed in place), and in `finish`. Disabled, it
    checksums every write in place, exactly as before. An allocation failure
    only disables it. Its `Debug` prints the staged length, never the bytes.
  * `ZipArchiveWriter` gains `reusable_crc_stage`, lent to each owned entry
    (`start_file_owned*`, the path every streaming Office writer takes) and
    returned by `ZipOwnedEntryWriter::finish`, as change
    [0618](0618-zip-writer-deflate-state-reuse.md) does for the Deflate state.
    Borrowed entries (`ZipDataWriterConfig::wrap`, used for bulk member
    copies) keep the stage disabled. `ZipDataWriter::finish` delegates to a
    private `finish_with_crc_stage`.
* `crates/litchi-docx/src/streaming.rs`
  * `for_each_escaped_chunk` replaces the character loop of
    `emit_escaped_text`: runs of bytes that need no entity are copied whole,
    and each chunk handed to the part is exactly the chunk the old loop
    flushed.
  * `scan_plain_text` replaces the character loop of `encoded_text_size`: an
    eight-byte SWAR test (`is_plain_ascii_word`) measures eight plain ASCII
    characters at once (`0x20..=0x7F` other than `&`, `<`, `>`), taken only
    when the step cannot cross the 64-character Work checkpoint; everything
    else goes one character at a time as before.
  * The writer's fields other than its `ExecutionContext` and scratch
    reservation move into a private `DocumentState`, so `write_text` holds
    its `InputBytes` reservation as a `ScopedReservation` borrowed from the
    context while the state emits. The public API is unchanged. The move
    changes the writer's field drop order: the part, with the caller's sink,
    still drops first and the scratch reservation is still released after it
    and after the writer's context handle, but that handle now drops after
    the other state fields instead of second, and `poison` before the scratch
    reservation instead of last. Nothing observes the difference: the only
    `Drop` impls involved are the two reservations' releases, and the context
    and the poison reason only release shared handles and error values.
* `crates/litchi-xlsx/src/streaming.rs`: `write_row` holds its row's
  `Objects` reservation as a `ScopedReservation`. The poisoning helpers borrow
  only the fields they poison (private `PoisonTarget` and `poison_target!`),
  so on a write failure the part is still poisoned and finished before the
  reservation is released, in the original order.

**Breaking changes: none.** The three additions are additive. Only private
representation and `Debug` output of `Reservation`, `ZipDataWriter`,
`ZipArchiveWriter` and the two writers change; `ZipDataWriter::finish` keeps
its signature. The one visible resource difference is one 4,096-byte
allocation per streaming archive (see *Allocations*).

## The evidence that motivated it

The coordinator's frame-pointer profile of the base's timed regions (base
`009d515bef`, CPU 28, summaries only) attributed the large DOCX stream (92 ms)
to budget accounting (about 25–30%: the atomic update in `consume` about
21%, `drop<Node>` 6% and `clone<Node>` 1.6% from `reserve`), CRC-32 (11%
self, 21.7% inclusive, almost all crc32fast's byte-at-a-time `update_slow`),
text escaping (28% inclusive) and Deflate (15.5%); the large XLSX stream to
Deflate (76%) and budget (14%); the large PPTX stream to Deflate (30%, half of
it the per-member sync flush), per-member zlib-rs resets (memset 10%),
allocation, hashing and part-name validation.

A flat profile of this record's before binary (30 samples of the large DOCX
case, CPU 28; untimed corpus construction and reopen gates are a few percent
of the samples; `profiles/docx-base.top.txt`) agrees:
`ExecutionContext::consume` 18.6%, `crc32fast::baseline::update_fast_16`
17.3% plus `Hasher::update` 1.6% and `new_with_initial` 0.5%,
`deflate_medium` 7.2%, `emit_escaped_text` 7.1%, `write_text` (with the
inlined text scan) 4.6%, `Reservation::commit` 4.1%, `Budget::reserve` 2.6%,
`append_character` 2.3%.

Why the CRC-32 was scalar: the harness builds with its own lock file
(`tools/perf-baseline/Cargo.lock`), which resolves crc32fast 1.5.0; its
carry-less-multiply path starts at 128 bytes, and every DOCX write (5, 31,
about 56, 12 and 6 bytes per paragraph) is shorter, so each byte took a table
lookup, and writes under 64 bytes the one-byte loop. The workspace lock
resolves crc32fast 1.5.1, whose SIMD path starts at 16 bytes. Staging serves
both. A fold triggered by a full stage covers more than 3 KiB: a write of at
most 1,024 bytes overflows the 4,096-byte stage only when it already holds
more than 3,072. A write longer than 1 KiB is checksummed in place. Only the
fold made just before such a longer write, and the one at `finish`, can be
short, as short as one byte (one byte staged, then a 2 KiB write, folds one
byte). The DOCX document member's writes are all shorter than 1 KiB, so every
fold but its last is triggered by a full stage. On the exact DOCX payload
(14,418,099 bytes in 655,362 writes) the probe measures 26.28 ms for
per-write CRC-32 against 2.25 ms staged at 4 KiB (0.80 ms for one buffer)
(`probe/crc_deflate_probe-output.txt`). As supporting evidence outside this
record's measured matrix, the coordinator reports that the independent review
ran the large DOCX case with the workspace's crc32fast 1.5.1 and saw about
85–87 ms before and 52–54 ms after, with the same output digest.

What each budget operation costs. The large DOCX writes 131,072 paragraphs,
each with seven `consume` charges and one input-byte `reserve`/`commit`. A
locked read-modify-write that follows the emission's stores (the Deflate
window copy and hash-table inserts) waits for them, which is why
`consume` costs roughly 20 ns a call in the profile (18.6% of the samples over
917,504 charges per iteration). The seven charges are each one
atomic update that exact, shared-budget semantics require, and the change
leaves them alone. The reservation's two or three reference-count updates
were pure overhead. A throwaway experiment (not committed; before the scoped
reservation existed) replaced `write_text`'s owned reservation with a plain
`consume` of the same amount — the same work on the success path minus the
reference counting: six processes each, the large DOCX median p50 went from
49.157 ms to 46.675 ms (−5.0%;
`experiments/scoped-reservation-upper-bound/`). The borrowed reservation
takes that saving with identical semantics on every path.

## Coalescing before the compressor is not byte-transparent

The task's second direction assumed that zlib's `Z_NO_FLUSH` output does not
depend on how the input is split into calls, so a staging buffer could feed
CRC-32 and Deflate large chunks. For the codec the harness and the workspace
use (flate2 1.1.9 over zlib-rs 0.6.7, level 6, whose `deflate_medium` path the
profiles show), it does depend on it. The probe
(`probe/crc_deflate_probe/`, CPU 28) rebuilds the large corpus's
`word/document.xml` and compresses it through the same one-codec-call-per-write
loop as `ReusableDeflateState::write_to`, with the sync flush and finish of
`ZipDataWriter::finish`. Its stream for the writer's split is byte-identical
to the real writer's member on both legs (376,031 bytes, SHA-256
`a29c4912…`; `probe/model-vs-writer.txt`). Re-splitting the same payload:

| input split | compressed bytes | identical to the writer's stream | inflates to the payload |
| --- | ---: | --- | --- |
| the writer's 655,362 writes | 376,031 | — | yes |
| 1-byte writes | 375,431 | no | yes |
| 7-byte writes | 375,545 | no | yes |
| 64-byte writes | 375,126 | no | yes |
| 4 KiB writes | 375,276 | no | yes |
| 16 KiB writes | 375,255 | no | yes |
| 64 KiB writes | 375,250 | no | yes |
| one call | 375,250 | no | yes |

Coalescing compressor input would therefore change every streaming package's
bytes, and the task forbids that. Only the CRC-32, a function of the byte
sequence alone, is staged. For the record, the probe prices what coalescing
would have saved on this payload: the per-write Deflate loop takes 21.81 ms
against 11.27 ms in 16 KiB chunks. It stays unclaimed (see *What remains*).

The same fact constrains the DOCX escaping rewrite. The 64-byte escaping
scratch decides which writes the Deflate stream sees, so the new code must
reproduce the old chunk boundaries exactly, not merely the concatenated
bytes.

## Why the bytes, the refusals and the timing of every error cannot move

* **CRC-32.** The stage folds accepted bytes into the CRC in their original
  order; CRC-32 over a concatenation equals CRC-32 continued over its parts.
  The stage is folded before a longer write is checksummed and in `finish`,
  before the data descriptor is formed, and the running CRC is observable
  nowhere else. What reaches the compressor, the output buffer and the sink,
  and when, is untouched.
* **Budget.** Every charge is still one atomic check-and-add per level,
  innermost first, with the same refusal value and rollback. Only how the code
  reaches each level (by reference instead of by a cloned handle) and how a
  reservation remembers its node change. A scoped reservation charges and
  releases through the same functions as an owned one; being borrowed, it
  cannot outlive its budget. In both writers the reservation is taken and
  committed at the same point of the same call, and on an error path it is
  still released after the writer is poisoned (DOCX: at the `?` after the
  failed emission; XLSX: after `poison_target!(self).io(error)` finishes the
  part), as before.
* **Escaping.** `for_each_escaped_chunk` emits a chunk exactly when the next
  character's escape does not fit the 64-byte scratch, the old loop's rule; a
  run of plain bytes is cut at its longest whole-character prefix that fits,
  and at most `room + 1` bytes are inspected per step, so the scan is linear
  in the text. `scan_plain_text` validates and counts the same characters,
  charges Work in the same amounts (64, then the remainder) after the same
  characters, and reports an invalid character before any later checkpoint
  and a refused checkpoint before any later character. The cancellation check
  before each chunk and every error mapping are unchanged.

## Measured

Host AMD EPYC 9R45, `Linux 7.0.0-1012-aws x86_64`, shared with other agents;
every measured process pinned to CPU 28 with `taskset`; no build of this
record ran during a measurement. Harness `tools/perf-baseline`, unchanged by
this record, with its deterministic generators. Both legs are built by the
identical command from their own trees: before from a detached worktree at
`6d989cad63` (`ffd2e348…`), after from the branch at `1167398a3e`
(`8d984f37…`); `binaries.txt` has the commands and every digest, and each
report records its revision, a clean tree and CPU 28.

### Timing

Four rounds; in each, every case ran before, after, after, before, so eight
processes per leg, paired within the round (slot 1 with 2, slot 4 with 3).
The table shows the medians of the process p50s and p95s, the median paired
p50 change and a percentile bootstrap 95% interval over the eight paired
changes (20,000 resamples, seed 752). Raw reports: `timing/abba-raw.tar.gz`;
summary: `timing/summary.json`.

| case, corpus | processes × samples | before p50 ms | after p50 ms | paired p50 change | 95% CI | p95 ms, before → after | output |
| --- | --- | ---: | ---: | ---: | --- | --- | --- |
| `docx_streaming_create` large | 8+8 × 20 | 91.474 | 48.596 | **−46.66%** | [−47.25%, −45.24%] | 92.236 → 49.287 | identical |
| `docx_streaming_create` medium | 8+8 × 40 | 5.754 | 3.094 | **−46.30%** | [−46.75%, −46.16%] | 5.771 → 3.122 | identical |
| `xlsx_streaming_create` large | 8+8 × 15 | 161.120 | 158.310 | −1.94% | [−2.67%, −1.19%] | 162.605 → 158.996 | identical |
| `xlsx_streaming_create` medium | 8+8 × 30 | 10.794 | 10.454 | −3.34% | [−4.58%, −2.32%] | 10.853 → 10.525 | identical |
| `pptx_streaming_create` large | 8+8 × 15 | 184.676 | 181.510 | −1.45% | [−2.49%, +0.30%] | 188.126 → 185.300 | identical |
| `pptx_streaming_create` medium | 8+8 × 30 | 6.420 | 6.298 | −1.86% | [−3.34%, −0.65%] | 6.466 → 6.356 | identical |
| control: `docx_semantic_one_edit_save` large | 8+8 × 40 | 8.446 | 8.448 | +0.10% | [−1.87%, +1.33%] | 8.546 → 8.597 | no digest reported |
| control: `xlsx_ordinary_save_lifecycle` | 8+8 × 20 | 14.271 | 14.126 | −1.15% | [−2.79%, −0.24%] | 14.793 → 14.333 | identical |
| control: `cfb_open_stream_mini_shared_bulk` many-small | 8+8 × 40 | 0.079 | 0.081 | +2.52% | [+1.81%, +3.47%] | 0.086 → 0.088 | no digest reported |
| control: `cfb_open_stream_mini_shared_bulk` wide-root | 8+8 × 40 | 0.650 | 0.670 | +3.06% | [+2.90%, +3.20%] | 0.666 → 0.690 | no digest reported |

"Identical" means one output SHA-256 per case across all sixteen processes:
`adf297d1…` (DOCX large, also the base's reference digest), `62ad295c…`,
`cf6965fc…`, `c8ac0ef5…`, `c7b08da6…`, `1f33f8b2…` and `20335c44…`.
The CFB and semantic-edit cases publish no digest in their reports. The
controls exercise the shared budget (the CFB shared bulk read reserves and
consumes through the execution context) without the writers.

### Instructions and cycles

User-space counters per timed iteration, by differencing a low- and a
high-sample run of the same binary (`scripts/perf_delta.py`), which cancels
corpus construction and the reopen gates. Two rounds of before, after, after,
before; medians of four measurements per leg, median paired change
(`counters/summary.json`):

| case | instructions before → after | change | cycles before → after | change |
| --- | --- | ---: | --- | ---: |
| DOCX large | 1,589.30M → 1,152.80M | −27.46% | 409.74M → 219.54M | −46.33% |
| DOCX medium | 100.24M → 72.90M | −27.28% | 25.92M → 14.65M | −43.29% |
| XLSX large | 2,412.17M → 2,384.07M | −1.16% | 727.32M → 714.76M | −0.54% |
| XLSX medium | 153.31M → 151.23M | −1.36% | 48.42M → 46.68M | −2.53% |
| PPTX large | 2,894.67M → 2,880.01M | −0.67% | not interpretable | — |
| PPTX medium | 98.91M → 98.39M | −0.53% | 29.14M → 28.43M | −1.88% |
| control: semantic edit | 435.00M → 434.87M | −0.02% | 93.31M → 95.64M | +3.36% |
| control: ordinary save | 73.48M → 73.45M | −0.02% | 27.88M → 28.01M | −0.90% |
| control: CFB shared bulk | 15.87M → 15.87M | −0.00% | 3.92M → 3.77M | −3.58% |

Each PPTX large process spends about 120 G cycles building and reopening its
16,421-member corpus outside the timer, so its cycle difference is lost in
that noise; its instructions are deterministic. The semantic-edit control's
+3.36% cycles with unchanged instructions matches its flat wall clock
(+0.10%) within the noise of these two-run differences.

### The CFB control's +3%: code placement, not work

The shared CFB bulk read slows by 2.52% and 3.06% although it executes the
same instructions (15,875,774 per iteration before, 15,875,031 after). Three
checks attribute it to code placement:

* **A/A.** The unchanged base built a second time by the identical command
  at a second path (`633a426d…`), measured against the before binary in the
  same ABBA design, moves the wide-root corpus +1.47% [+0.10%, +3.16%], and
  that campaign has eight adverse over-5% comparisons of its own across the
  CFB and ordinary-save controls; the other controls' medians stay within
  ±0.4% (`layout-control/aa-summary.json`). Two builds of identical
  code already differ by up to about 3% on this 0.65 ms case.
* **Counters.** Per iteration, the before, layout and after builds execute
  15,875,774, 15,875,791 and 15,875,031 instructions, with 5,782, 6,016 and
  6,157 branch misses and 255,949, 297,264 and 267,697 front-end stall cycles
  (medians of four differenced measurements each; directory `frontend/` in
  `layout-control/aa-raw.tar.gz`).
* **Where the cycles go.** In profiles of 1,500 harness samples per build,
  the case's time sits in unchanged `litchi-cfb` code whose share moves with
  placement:
  `OleFile::load_directory` is 3.09% of samples before, 5.57% in the layout
  build of identical source and 6.43% after; `try_push` 3.04%, 3.12% and
  4.47%. No budget function appears among its top symbols
  (`layout-control/cfb-*.top.txt`). Its handful of budget calls per iteration
  now take fewer atomic operations than before.

The A/A median (+1.47%) is about half of the shift, though its interval
reaches +3.16%. The shift is therefore reported rather than dismissed: it is
the largest adverse median of the campaign.

### Regressions and every over-5% flag

No streaming case regresses. The largest adverse median paired change is the
CFB control's +3.06% above. Of the 62 paired comparisons over 5% at p50, p95
or mean (`timing/flags.json`), 51 are favourable: all 48 of the two DOCX
streaming cases and three of the ordinary-save control. The 11 adverse ones
are on controls: eight on the CFB bulk read — five on the 0.08 ms corpus
(four p95s, +5.6% to +12.2%, and one mean, +5.7%) and three on one wide-root
pair in round 1 (+5.8% p50, +5.6% p95, +5.8% mean) — and three on
`xlsx_ordinary_save_lifecycle` (round 1 p95 +5.1%, round 4 p95 +35.9% and
mean +8.7%, one slow tail process). The A/A campaign of identical code
produces the same kinds of flags on the same two controls (eight adverse;
`layout-control/aa-summary.json`).

### Allocations

Allocator counters per timed iteration from the allocation build
(`alloc/summary.json`; one process per leg and case, medians over the
samples):

| case | allocation calls | allocated bytes | region peak live bytes |
| --- | --- | --- | --- |
| DOCX tiny / medium / large | 104 / 8,232 / 131,112 → +1 each | +4,096 each | +4,087 each |
| XLSX tiny / medium / large | 135 / 8,263 / 131,143 → +1 each | +4,096 each | +4,087 each |
| PPTX tiny / medium / large | 851 / 9,050 / 278,154 → +1 each | +4,096 each | +4,087 each |
| control: ordinary save | 29,848 → 29,848 | unchanged | 14,971,825 → 14,971,816 |

The one extra allocation per streaming archive is the CRC stage (4 KiB,
allocated at the first short write and lent to every later member). The
reservations allocate nothing, before or after: a four-level `SmallVec` stayed
inline. The CFB and semantic-edit cases report no allocator counters.

## Correctness evidence beyond the tests

* **Differential probe, base against candidate.** One source file
  (`differential/diffprobe_main.rs`) built twice with the same lock file,
  path dependencies into each tree, drives the three streaming writers
  through 28,266 scenarios and prints every call's result or error value, the
  reported `written`/`output_bytes` progress, every budget counter of every
  level after every call, and each output's SHA-256: every value of every
  budget dimension (Work, Objects, InputBytes, OutputBytes, Memory) from zero
  to one past the reference run's total — for Memory, past the writers' fixed
  reservations of 64 and 4,096 bytes — for the DOCX and XLSX writers, each
  both on a single-level budget and as the parent of a child; sweeps of each
  writer-local byte limit, including the PPTX output limit; cancellation
  before each of the DOCX script's 56 calls and inside the sink at 231 byte
  positions; sinks that fail at every 11th (DOCX), 17th (XLSX) and 173rd
  (PPTX) byte, with short writes (DOCX, XLSX) and zero progress (DOCX); and
  `\t`, U+FFFE and U+0001 at every third position of a 141-character text
  under seven Work limits placed around its checkpoints. The transcripts are
  byte-identical (`d15ea98b…`, 39 MB; per-category counts in
  `differential/outcomes.txt`: every refusal kind at every call type,
  including 2,550 Work refusals inside `write_text`, 813 invalid inputs and
  128 Memory refusals at construction).
* **Output identity.** Every harness process of both legs, in every case with
  a digest, published the same package; the real large DOCX written by each
  leg's probe is `adf297d1…` with the same `document.xml` stream.
* **The existing owned-versus-borrowed differential** (`deflate_reuse.rs`)
  compares complete ZIP bytes, CRC fields included, between the owned path
  (stage enabled) and the borrowed path (stage disabled); a new feed plan
  crosses the stage's thresholds.

## What is not claimed

No claim is registered. The numbers are scoped to the named synthetic corpora,
this host and CPU 28, and the harness lock's crc32fast 1.5.0. They do not
establish:

* the saving in production builds with the workspace lock's crc32fast 1.5.1,
  whose SIMD path starts at 16 bytes, so writes of 16–127 bytes were already
  cheaper there: this record did not measure it; the review's informal
  observation above (about 85–87 → 52–54 ms) is supporting evidence only;
* behaviour with shared or deep budget hierarchies under concurrency: the
  harness contexts are single-level roots, and the reservation change was
  exercised concurrently only by the unit tests;
* RSS, cold cache, other sinks, other producers or other platforms;
* anything about the remaining per-call costs listed below.

## What remains

* **Per-call budget charges.** The seven `consume` charges per DOCX paragraph
  and five per XLSX row are each one atomic update; they are now the largest
  item of the DOCX profile (26% of samples, `profiles/docx-after.top.txt`).
  Fewer of them requires a semantic decision — charging Objects and Work
  together, or leasing budget to a writer — that changes what other holders
  of a shared budget observe. Removing one locked operation also moves part
  of the store-drain wait to the next: after this change `finish_run`, whose
  charge is the first locked operation after the text's emission, takes 4.7%
  of the samples, up from under 0.5%, while `Reservation::commit`'s 4.1% is
  gone.
* **Deflate call overhead.** About 10 ms of the 48.6 ms large DOCX iteration,
  by the probe. Recovering it by coalescing would change the published bytes
  (still deterministic, but different from today's); that is a decision for
  the owner, not something this record could do.
* **PPTX per-member costs**: the zlib-rs reset of each member's hash table,
  the per-member sync flush and final block, part-name validation and hashing.
* XLSX cell escaping (`push_escaped`, 2.3%) is already run-batched.

## Authority

ADR 0005 (every operation charges a hierarchical budget; limits stay finite
and exact; bounded streaming — the budget semantics are unchanged and the new
stage is a fixed 4 KiB transport buffer outside the semantic accounting, like
the Deflate output buffer), ADR 0006 (deterministic serialization — no byte
moves), ADR 0001 (the facade stays panic-free: the new code uses checked
access and fails closed), ADR 0010, 0011 and 0024 (the stage lives in the
archive owner; no archive type crosses into format crates; no new dependency
edge). Owner decisions of change
[0652](0652-owner-decisions-for-the-third-wave.md): standing trade-off 2 —
the exact route was taken over coalescing, so every refusal, limit,
cancellation point and output byte is kept — and trade-off 3 — plain ASCII
text takes the eight-byte step, and any other input takes the unchanged
per-character path. Decision 8's allowance to move error timing was not
needed and not used.

## Verification

All gates pass at `1167398a3e` (`results/change-0752/gates.txt`):
`cargo fmt --all --check`; `cargo check --all-targets` of the four touched
crates and every in-scope dependent (29 packages; three pre-existing warnings
in the untouched `crates/litchi/tests/unexpected_format.rs`); warning-denied
Clippy on the library and all targets of the touched crates; warning-denied
rustdoc; `cargo test` of `litchi-core` (222), `soapberry-zip` (643),
`litchi-docx` (1,530), `litchi-xlsx` (1,420) and the 24 other in-scope
crates that depend on them, the facade with `doc,docx,ppt,pptx,xls,xlsx,xlsb,odt` (382) and the
unchanged harness (553); crate boundaries (64 packages, 241 declarations, 11
explicit debts); the non-iWork gate; the structural claims check (10 claims).

New tests: `litchi-core` —
`limits_refuse_at_exactly_one_over_at_the_innermost_refusing_level`
(N−1/N/N+1 through `consume` and `reserve` on a three-level hierarchy whose
middle level is tightest), `scoped_reservations_charge_commit_and_release_exactly_as_owned_ones`,
`partial_and_zero_commits_release_every_level`,
`a_reservation_holds_one_handle_whatever_the_depth`,
`a_reservation_outlives_the_budget_handles_that_made_it`,
`deep_hierarchies_roll_back_exactly` (the former spill test), and the
cancellation test extended to `reserve_scoped`. `soapberry-zip` —
`crc_stage_matches_a_one_pass_checksum_at_every_boundary`,
`crc_stage_continues_from_a_custom_initial_value`,
`owned_entries_checksum_short_writes_exactly_and_reuse_one_stage`, and the
streaming-shaped plan in `deflate_reuse.rs`. `litchi-docx` —
`escaped_chunks_match_the_character_loop` and
`text_scan_matches_the_character_loop_at_every_work_limit` (the old loops
kept as oracles: 419 texts for the chunking and 613 for the scan, each at ten
Work limits, hand-picked boundary cases plus generated texts over every
character class, refused characters included), and
`plain_ascii_word_matches_a_byte_by_byte_test` (every byte value at every
position, and 200,000 boundary-biased words).

## Review follow-up

The coordinator relayed an independent review that recommended merging with
no blocker. It reproduced the 28,266-line transcript and ran its own
differentials: 1.6 million budget steps over owned, scoped and base
reservations, a 16-thread stress test (200,000 operations per thread, three
seeds), 7,000 archives of random writes (14,780 CRCs verified), and 122,000
texts (2.73 million runs) through the escaper. All matched exactly. Its four
nits are folded in without rewriting history:

* The fold-size sentence above was narrowed: only folds triggered by a full
  stage are guaranteed to exceed 3 KiB. The review's crc32fast 1.5.1
  observation is added as supporting evidence.
* `crc_stage_matches_a_one_pass_checksum_at_every_boundary` had an assertion
  that was vacuous for an enabled stage. It now asserts that a disabled stage
  never holds bytes, and, for an enabled one, that an empty write changes
  nothing, a short write leaves the stage ending with exactly its bytes, and a
  long write leaves it empty.
* New `concurrent_owned_and_scoped_reservations_keep_a_shared_hierarchy_exact`
  in `litchi-core`: eight threads on a six-node, three-level tree (two
  sibling leaves under one middle node, one leaf under the other; two
  workers on each leaf, one on a middle node, one on the root) run fixed
  pseudo-random mixes of owned and
  scoped reservations of Memory and Work, drops, zero commits, over-commits,
  partial Work commits and Work consumption. Every refusal must name a level
  of its chain with that level's limit. A monitor thread reads every level
  throughout and must never see one above its limit. At the end, every
  level's Memory must be zero and its Work exactly its subtree's granted sum.
  The seeds ask for more Work through `consume` alone than the root allows,
  so at least one refusal is certain in any interleaving. It takes about
  10 ms in a debug build. As a mutation check, rolling back one level too few
  on a refusal failed it in 10 of 10 runs (22 units of Memory left at
  `left`), and a check-after-add charge failed it in 10 of 10 runs (the
  monitor saw `left_b` over its limit); both mutations were reverted and the
  file compared equal.
* The drop-order sentence under *What was changed* was added.

The follow-up gates are appended to `results/change-0752/gates.txt`.

## Cleanup

Binary and probe digests are in `binaries.txt`. After the evidence was
copied, the target directories (`targets/0752`, `targets/0752-before`,
`targets/0752-layout`, the three probe targets and the earlier experiment
target), the detached before and layout worktrees and the scratch directory,
including every `perf.data` and the 39 MB transcripts, are removed; the
branch worktree is kept (`results/change-0752/cleanup.json`).
