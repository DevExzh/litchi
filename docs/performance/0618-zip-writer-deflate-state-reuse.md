# 0618: one Deflate compressor per authored save instead of one per member

Status: retained, implemented — with one of the two proposed mechanisms
measured and rejected. `performance_claim: none` — no claim-registry entry is
created by this wave; the paired medians, native cycle counts and instruction
counts below are reported as evidence, not registered as claims.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred
until that goal completes; iWork is excluded.

This record implements item **SAVE-4** of
[0587](0587-remaining-opportunity-survey.md) at the two sites change
[0607](0607-pptx-authored-slide-regeneration-design.md) named; 0607 landed on
the branch after this record's base, so its file resolves once this work is
merged. Base `8fe9efa55728cf8b9592934f9e9de6f5a66fdc1b`; branch
`perf/0618-zip-preservation-writer-deflate-reuse`.

## What was changed

Three mechanisms in `crates/soapberry-zip`, all of them work elimination.
Nothing about which bytes are published changed, on any input.

**1. The reusable Deflate state is no longer private to the owned writer.**
Change 0476 gave `ZipArchiveWriter` an `OwnedDeflateState` — one `flate2`
`Compress` plus one 32 KiB output buffer, driven by hand so that its write,
flush and finish boundaries reproduce `flate2::write::DeflateEncoder`'s exactly
— but only `start_file_owned` could reach it. It is renamed
`ReusableDeflateState`, made `pub(crate)`, and its `write_to` / `flush_to` /
`finish_to` helpers are generalized from `OwnedCompressedEntry<W>` to any
`W: Write`. A new `pub(crate) ReusedDeflateEncoder<'state, W>` wraps a borrowed
state behind the same `Write` + `finish` surface `DeflateEncoder` offers, so the
call sites keep their shape. `ZipArchiveWriter` gains `take_reusable_deflate`
and `restore_reusable_deflate`; the owned route now goes through them.

**2. The reset moved from the end of a member to the start of the next one.**
`ReusableDeflateState` carries a `used` flag and a `begin_member` method that
resets only when a previous member left a stream in the compressor. A save with
a single Deflate member therefore pays exactly the one construction it paid
before and no reset at all. The 32 KiB output buffer moved from an inline
`[u8; 32 * 1024]` to a `Box<[u8]>`, so the struct is a few words wide and moving
it into its box no longer zeroes 32 KiB on the stack and copies it to the heap.

**3. Both Office write sites drive the compressor instead of constructing one.**
`StreamingArchiveWriter` (`src/office.rs`) — the writer behind every authored
XLSX, DOCX and PPTX publication, through `litchi-opc`'s `phys_pkg.rs` — takes
the state from its archive before opening a Deflate member and returns it only
after that member's final Deflate output succeeded, at all three of its Deflate
sites (`write_deflated_with_accounting`, `write_deflated_sized_with_accounting`
and the reader-fed `write_reader_with_accounting`). One archive now constructs
one compressor. The preservation writer (`src/preserve.rs`, `generated_entry`)
keeps **one compressor per regenerated member**, constructed and freed inside
it, but builds it as a `ReusableDeflateState` driven by `ReusedDeflateEncoder`
rather than a `flate2::write::DeflateEncoder`; that alone removes flate2's
per-codec-call zeroing of its output vector's spare capacity.

What was **not** changed: the compression level, the strategy, the framing, the
data descriptors, the ZIP64 decisions, the limits, the accounting, the public
API, and the one-entry mini-archive the preservation writer builds to derive a
regenerated member's framing (see *Measured*).

## Why it is sound

**Reuse is byte-transparent by construction.** `Compress::reset` is zlib's
`deflateReset`: it restores the level, strategy and window the constructor
selected and clears the stream, so the member that follows emits exactly what a
freshly constructed encoder emits. `ReusedDeflateEncoder` keeps 0476's
hand-driven boundaries unchanged — one codec call per `Write::write`, the
sync-flush ordering of `zio::Writer::flush`, and the drain-then-`Finish` loop of
`zio::Writer::finish` — so the *sequence* of codec calls a member sees is the
one flate2 would have produced. The output buffer's 32 KiB size is unchanged,
which matters beyond speed: `LimitedEntryWriter` evaluates the compressed-size
budget per `write` call, so a different drain granularity could move when a
limit is refused.

**A failed member cannot contaminate the next one.** The state leaves its owner
for the lifetime of one member. On success it is returned; on any failure it is
dropped, and the next member constructs a fresh one — the discipline 0476
established for the owned writer, now applied at every site. `ReusedDeflateEncoder`
also reproduces flate2's `Drop`, which finishes an unfinished stream and
discards the result, so a member abandoned part-way emits the bytes it emitted
before and the accounting that follows is unchanged.

**Error identity is untouched.** No limit moved, no refusal moved, and the
progress-failure text is deliberately left as it was. The only new error is a
defensive `InvalidInput` for a writer-internal state that the writer itself
always prepares; it is unreachable from any caller.

**ADR reading.** ADR 0006 (lossless preservation) is the binding constraint:
published bytes must not change. That is not argued, it is measured — 2,352
corpus rows, 422 editor rows and every authored-save digest are byte-identical
between legs. ADR 0005's bounded resources are unaffected: the compressor that
existed per member now exists per archive at the streaming writer and still per
member at the preservation writer, so no resource ceiling rises. ADR 0011 is
unaffected: `ReusableDeflateState` and `ReusedDeflateEncoder` are `pub(crate)`
and no archive type, lock or executor becomes visible. No `unsafe` is added, no
defence is weakened, and no new ambient I/O or global pool appears.

## Measured

Host: AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws; rustc 1.95.0; valgrind
3.26.0; every measured process `taskset -c 12`; seven other agents active, run
window load average 20-25. Build `--release --locked` on both legs. Scenarios:
the streaming writer through an authored PPTX save (change 0607's probe, 161
output members at 50 slides), and the preservation writer through
`ConditionalFormattingSamples.xlsx` (132 members, 654,688 bytes) with 1, 8 and
40 `.rels` members regenerated.

### Deflate-compressor constructions per save (callgrind call counts)

| save | members | before | after |
| --- | ---: | ---: | ---: |
| authored PPTX, 50 slides | 161 | 161 | **1** |
| authored PPTX, 1 slide | 39 | 39 | **1** |
| preservation, 40 regenerated members | 40 | 40 | 40 |
| preservation, 1 regenerated member | 1 | 1 | 1 |

### Instruction counts (callgrind isolation pairs) — an upper bound

Callgrind counts a `rep stos` once per byte, and a compressor's ~300 KiB state
zeroing is exactly that, so these deltas over-price what was removed. They rank
the work; the native numbers below price it.

| scenario | before Ir | after Ir | delta |
| --- | ---: | ---: | ---: |
| authored PPTX, 50 slides | 159,989,232 | 91,439,843 | **-42.85%** |
| authored PPTX, 1 slide | 42,391,193 | 26,010,098 | **-38.64%** |
| preservation, 40 regenerated | 38,527,311 | 33,270,402 | -13.64% |
| preservation, 8 regenerated | 9,671,970 | 8,618,644 | -10.89% |
| preservation, 1 regenerated | 3,245,905 | 3,119,576 | -3.89% |

### The native price (`perf stat -x, -r 5`, same isolation pairs)

| scenario | metric | before | after | delta |
| --- | --- | ---: | ---: | ---: |
| authored PPTX, 50 slides | cycles | 14,512,118 | 12,735,482 | **-12.24%** |
| | instructions | 48,071,748 | 45,920,075 | -4.48% |
| | minor faults | 11.0 | 12.2 | +1.2 |
| authored PPTX, 1 slide (39 members) | cycles | 5,063,727 | 4,358,787 | **-13.92%** |
| | instructions | 15,757,367 | 14,737,266 | -6.47% |
| | minor faults | 94.9 | 0.0 | **-94.9** |
| preservation, 40 regenerated | cycles | 2,673,674 | 2,614,286 | -2.22% |
| | instructions | 9,559,247 | 9,415,296 | -1.51% |
| preservation, 8 regenerated | cycles | 832,243 | 816,657 | -1.87% |
| | instructions | 2,931,666 | 2,901,340 | -1.03% |
| preservation, 1 regenerated | cycles | 402,645 | 403,176 | +0.13% |
| | instructions | 1,371,972 | 1,368,759 | -0.23% |

The small-deck row is the clearest reading of the mechanism: a 39-member
authored save used to fault in and release 39 ~300 KiB compressor states, at a
net 95 minor page faults per save; it now constructs one and faults in none.

### Paired timing, A1 B1 B2 A2, both directions, with the floor

Authored PPTX save, 50 slides, 40 samples per leg-run (microseconds per save):

| leg | p50 | mean | p95 | p99 |
| --- | ---: | ---: | ---: | ---: |
| A (before, pooled n=80) | 3,299.46 | 3,306.92 | 3,355.94 | 3,555.69 |
| B (after, pooled n=80) | 2,899.69 | 2,900.39 | 2,939.20 | 2,951.08 |
| B - A | **-12.12%** | -12.29% | -12.42% | -17.00% |
| A - B | +13.79% | +14.02% | +14.18% | +20.49% |

A/A floor p50 -1.90%, p95 -3.33%, p99 -14.36%; B/B floor p50 -1.01%, p95
-1.26%, p99 -2.64%. The p50, mean and p95 deltas are several times the floor;
**the p99 delta is inside it and is not claimed**.

Preservation publish, 40 regenerated members, 2,000 samples per leg-run, two
windows (microseconds per publish):

| window | A p50 | B p50 | B - A p50 | A/A floor p50 | B/B floor p50 |
| --- | ---: | ---: | ---: | ---: | ---: |
| 1 | 608.48 | 596.40 | -1.99% | +0.39% | +0.57% |
| 2 | 601.69 | 588.43 | -2.20% | -0.66% | -0.85% |

Window 2's mean is -2.31%, p95 -3.43%, p99 -2.64%. Window 1's A1 leg carries a
tail outlier (p99 993 µs, A/A p99 floor -35.57%), so only window 2's tail is
readable. Both windows agree with the -2.22% cycle delta.

Preservation publish, **1** regenerated member: B - A p50 +0.39% (window 1) and
+1.24% (window 2), against A/A floors of -1.02% and -0.92%. This scenario is
inside the floor in both directions; **no change is claimed for it**, which is
what the brief required.

### Rejected: one compressor shared across a plan's regenerated members

The brief's second half — carry one compressor across the members of a
preservation save — was implemented, measured and removed. It worked as
designed on the counts (40 constructions per publish became 1; callgrind Ir
-42.74%) and lost badly on the machine:

| scenario | metric | before | pooled | delta |
| --- | --- | ---: | ---: | ---: |
| preservation, 40 regenerated | cycles | 2,684,598 | 3,063,117 | **+14.1%** |
| | instructions | 9,564,837 | 10,514,880 | **+9.9%** |
| | minor faults per publish | 0 | **271** | +271 |

Its paired timing was bimodal: the after leg settled at either 531 µs or 665 µs
against a stable 598 µs before, with a one-way transition mid-run in one leg-run
(`timing/` is the post-removal harness; the bimodal series is described here
rather than retained). Mechanism: `prepare` retains one `Arc<Vec<u8>>`
mini-archive buffer per regenerated member for the publication pass, so a
~300 KiB compressor held across those members is freed *under* them and glibc
cannot recycle it; the next publish takes it from `mmap` and returns it, at 271
minor faults per publish. Constructing and freeing it inside each member —
which is what the before leg did and what this change keeps — lets the allocator
reuse the same chunk. The pooled variant is not retained in any form; its
replacement is the member-local state described above, which is a smaller but
real win on the same scenarios.

### Measured and deliberately not removed: the mini-archive round trip

`generated_entry` compresses a regenerated member into a one-entry
`ZipArchiveWriter`, then re-parses that mini-archive to locate its framing.
Removing the round trip was in scope only if the framing could be computed
directly with identical bytes. Priced first: `ZipArchive::from_slice`, the two
`next_entry` calls and the layout checks total 592,280 Ir over 960 regenerated
members — **617 Ir per member, 0.15%** of the 424,670 Ir a regenerated member
costs after this change (`get_entry` and `compressed_data_range` are inlined and
carry no attributable cost). Computing the framing directly would mean
reproducing the writer's descriptor width, ZIP64 promotion and central extra
field decisions inside `preserve.rs`, and dropping a structural check on a
preservation path, for 0.15%. Not taken.

## Correctness evidence

**Byte-identity oracles, every one of them empty-diff.**

| oracle | rows | result |
| --- | ---: | --- |
| `OpcPackage` open → mutate → publish, 336 OOXML fixtures x `noop`/`pkgrels`/`reblob`/`addrel` | 1,344 | identical, including 8 open-errors and 27 save-errors matched by typed-error text |
| the same 336 fixtures x `addrelN:2`/`addrelN:8`/`addrelall` (2 to 43 regenerated members) | 1,008 | identical |
| the real editor routes: `tabs` hide and `edit_cells` over every XLSX/XLSM, `append_plain_paragraph` over every DOCX/DOCM | 422 | identical: 110 published digests and 312 typed refusals |
| authored PPTX saves at 1, 2, 5, 10, 50 and 200 slides, twice each (stability and one-edit), SHA-256 of all four outputs | 48 digests | identical |
| 78 opened decks saved back | 78 | identical, 78 of 78 exact-source passthrough on both legs |

**Tests added.** Four in `crates/soapberry-zip/tests/deflate_reuse.rs`, beside
0476's owned-writer differentials:
`streaming_deflate_members_match_one_fresh_encoder_each` and
`streaming_sized_deflate_members_match_one_fresh_encoder_each` build a
five-member archive through `StreamingArchiveWriter` and require the complete
ZIP bytes to equal an archive built with one fresh `DeflateEncoder` per member;
`streaming_deflate_stream_members_match_one_fresh_encoder_each` does the same
for the reader-fed route against a reference fed the same chunks; and
`a_refused_sized_member_leaves_the_next_member_byte_identical` refuses a member
on the compressed-size limit and requires the member after it to be byte-identical
to the same archive written without the refusal. One in `src/preserve.rs`:
`a_regenerated_deflate_member_matches_a_fresh_flate2_encoder` compares a
regenerated member's local framing, payload and central record against the
one-entry mini archive a fresh `flate2` encoder produced, over six payload
sizes including empty and 64 KiB + 11.

**Gates.** `cargo fmt --all --check`, `cargo clippy -p soapberry-zip
--all-targets`, `cargo test -p soapberry-zip` (593 passed, 0 failed),
`cargo doc -p soapberry-zip --no-deps`, and the consumer suites `litchi-opc`
(675 passed), `litchi-xlsx`/`litchi-docx`/`litchi-pptx` (340 passed), and
`litchi-odt`/`litchi-odc`/`litchi-odf-common`/`litchi-iwa-archive` (352 passed)
— the other crates that write through `StreamingArchiveWriter`. All seven
sections exit 0; tails in `results/change-0618/gates.txt`.

**An incidental finding, unchanged by this work.** The reader-fed streaming
route and the whole-payload route produce slightly different Deflate bytes for
the same member, because zlib-rs's block decisions are not invariant under
regrouping of the input across codec calls. That was true before this change and
is true after it; the new reader-fed test pins it by feeding the reference the
same chunks.

## Validation preserved

Every limit, refusal and check on both write paths is untouched.
`LimitedEntryWriter` still charges the compressed-size budget per write and
still re-checks with `ensure_within_limit` after the final block;
`CompressedScratch` still refuses an oversized sized member before any archive
byte is emitted, leaving the writer usable — the new test proves the member
after such a refusal is byte-identical. The preservation writer still validates
the mini-archive it builds: entry count, directory and EOCD ordering, central
record framing and payload range, each with its own `unsupported` refusal. The
27 save-errors and 8 open-errors in the corpus run reproduce with identical
typed-error text, as do the 312 editor refusals. `verify_authored`, the
publication plan, the accounting counters and the ZIP64 promotion rules are not
touched.

## Limitations

* No claim is registered. Every number here is scoped to this host, this build,
  these fixtures and these scenarios.
* The p99 delta on the authored save is inside the A/A floor and is not claimed.
  The single-member preservation publish is inside the floor in both directions
  and no change is claimed for it.
* The regenerated members in the corpus scenarios are `.rels` parts of a few
  hundred bytes. No fixture in the corpus regenerates a large content part
  through the preservation writer, so the preservation numbers are the
  fixed-per-member cost, not a payload-size result.
* The streaming-writer result is measured on PPTX authoring. Authored DOCX and
  XLSX publication uses the same writer and the same code path, but was covered
  only by the byte-identity oracles and the test suites, not by its own timing.
* The page-fault mechanism behind the rejected pooled variant is glibc-malloc
  behaviour. A different allocator could rank that variant differently; nothing
  here claims it is unprofitable in general, only that it is unprofitable here.
* Wall-clock measurements were taken with seven other agents active on the host.
  Both timing windows carry their A/A floor; nothing is claimed below it.
* Cold-cache, physical-device, peak-RSS, allocation-profile and cross-platform
  behaviour were not measured.

## Retained evidence

[`results/change-0618/README.md`](results/change-0618/README.md) — the two
probes and their manifests, the callgrind and `perf stat` isolation pairs with
their runners, five paired-timing directories with per-sample data and the
A/A floor calculator, the five byte-identity oracles with their empty diffs, the
gate tails, `decision.json` and `log-sections.md`.
