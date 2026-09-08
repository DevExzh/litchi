# Change 0468: profile remaining dense XLSX commit/save work

`performance_claim: none; descriptive CPU attribution`

`claim_authorized: false`

This batch profiles the committed single-scan eager worksheet parser from
0467. No production or harness Rust changes. Clean revision `933cb6b80` has
the same 6,992 compiled-source inventory files and two compile-time fixtures
as production candidate `87733cf3b`. All 30 previously reviewed ADR file hashes
are unchanged. The fresh Rust 1.98.1 build has release debug level 1, forced
frame pointers and unwind tables; source, build and binary receipts bind it.

## Current observations

The ordinary dense one-percent commit/save workload edits 1,311 of 131,072
cells across two 256-by-256 sheets. Input SHA-256 is
`5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`.
All four lanes pass their output/readback oracle. Output stays 388,095 bytes,
37 sink calls, with a largest write of 65,536 bytes. Source opening, edit
staging and sink reservation precede the timer; commit and sequential write
are timed together. External tools include the complete process.

Two 30-sample, three-warmup uninstrumented runs on CPU 2 with one worker
measure p50 382.299942 / 382.735127 ms (0.114% same-build drift), with GNU
maximum RSS 109,728 / 109,732 KiB. These diagnostic-build observations do not
form a cross-build speedup claim or satisfy the registry's 500-sample minimum.

Whole-process perf stat reports 84,959,973,821 user cycles,
281,833,512,110 user instructions, 63,389,859,365 user branches,
84,005,741 user branch misses, 77,131,552 user cache misses and 556,388 page
faults. Every counter reports 100% running coverage. Derived IPC is 3.3173
and branch-miss rate 0.1325%. These totals include fixture generation,
expected-output construction, warmups, verification and teardown; they are
not counts per commit, per cell, or per retained timing sample.

The 50-sample/three-warmup, 499 Hz frame-pointer capture contains 15,164
accepted cycle samples and 132,347,593,651 weighted event periods. No samples
are reported lost; stacks containing unknown frames account for 0.592% of
whole-process weight. The exact `Edit::commit` ancestor covers 7,999 stacks
and 69,094,592,304 periods, or 52.21% of whole-process weight.

| Inclusive context | Share of exact commit context |
|---|---:|
| Eager worksheet Parser union | 40.01% |
| Snapshot scanner | 24.41% |
| Changed-XML compaction | 13.27% |
| Web publishing metadata reader | 10.33% |
| Snapshot cell address lookup | 4.93% |
| Shared unqualified attribute helper | 3.99% |
| Snapshot tag capture | 3.69% |
| Compaction start-tag emission | 3.20% |

These rows overlap. For example, address and tag processing sit inside the
snapshot scanner. The explicit CountingSink package writer accounts for
15.93% of whole-process sampled weight. Inclusive samples are not phase timers;
commit ancestors can also occur during untimed expected-output construction.
The separate `additional-summary.json` shows the web reader’s nearest
retained caller is commit in all 863 matching samples. Most nearest-child
weight is XML reader work (34.14% namespace-reader processing and 25.82%
event reading), supporting the full-pass explanation. The old 0466
percentages are historical context, not a numeric CPU baseline
with identical sampled denominators.

## Next measured experiment

First test removal of unconditional XML event ownership conversion in
`raw::compact::changed`. The current loop calls `into_owned()` on every event,
then fully handles or writes that event before reading the next one. The
borrowed input remains alive throughout. Keeping borrowed events should remove
copies and temporary allocations while retaining the same parse, normalize,
whitespace, entity, namespace, error and publication behavior. Consider merging
the two checked start-tag attribute traversals only as a separate reviewed step.

The explicit `Event::into_owned` marker accounts for 5.03% of the compaction
subtree (0.348% of whole-process weight); it is a small portion of the total.
Reader/namespace work and start-tag normalization remain. Additional inlined
copy/allocation effects cannot be quantified from that symbol alone, so a
small source change still needs allocation and latency measurement before
retention. If it is not material, proceed to the larger validated-pass reuse
work rather than treating this micro-optimization as the program outcome.

Eliminating the entire observed compaction subtree would bound sampled
commit-context speedup at approximately 1.153x. Event borrowing removes only
part of that work, so this is an upper-bound prioritization model, not an
end-to-end prediction. Snapshot scanning and eager parsing are larger remaining
paths; their ownership and validation obligations warrant separate work.
Adding a pre-sort monotonicity scan is not justified: Rust's unstable sort
already detects ordered input.

The next code change needs preservation/error differential tests, applicable
XLSX correctness gates, clean matched-build timing and resource comparisons,
and individual regression review. It must leave the bounded Store handoff at
4,096 cells / 1 MiB and retain publication validation. No SIMD or parallelism
change is indicated by this profile.

## Evidence and limits

Capture and postprocessing are serialized and their receipt chronology checks
pass. Instrumented maximum RSS (including perf) is not normal workload memory.
No new Heaptrack, allocation-count, copied-byte, native-producer, physical-cold,
range-source or worker-scaling result is claimed. Production correctness remains
bound to the unchanged 0467 source inventory and its passing 1,242-test suite;
this batch adds Python evidence tooling and a clean performance build.

See the [0468 bundle](../results/change-0468/README.md),
[source review](../results/change-0468/source-review.md),
[profile summary](../results/change-0468/summary-fp.json), and
[timing/counter observations](../results/change-0468/observations.json).
The broader non-iWork performance goal remains open.
