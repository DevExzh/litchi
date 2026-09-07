# 0463 — retain ODP publication audit proof reuse

0463 adds a private writer-origin proof to the ordinary ODP commit path. The
private ODP serializer records authored XML accounting around the existing
common `PackageWriter` audit path, only after each payload passes the strict
default audit. It binds the exact source and candidate archive owners and
conservatively bounds authored XML bytes and parts; `PackageWriter` itself is
unchanged. An eligible proof lets commit skip only the repeated final compactness scan; a
missing owner, limit mismatch, overflow, source or candidate `Arc` mismatch,
or later package replacement falls through to the existing validator.
Candidate reopen, semantic readback, source-reference precheck, media/domain
checks, no-op behavior and reversible patch construction remain in place.

The frozen A1/B1/B2/A2 matrix retains 24 reports and 720 samples. Normal p50
candidate-minus-baseline deltas are -10.0491% / -8.9678% for tiny,
-6.9642% / -7.3050% for medium, and -7.3657% / -6.8199% for large in R1/R2.
All normal and allocator p50 bootstrap intervals are below zero, the
predeclared 3% medium/large gate passes, and no adverse >5% elapsed or RSS
flag is present.

Allocator p50 deltas are -9.2947% / -8.8357% for tiny, -3.6356% / -3.6243%
for medium, and -4.0327% / -2.8299% for large. Allocated bytes fall by
2,087,682 / 11,072,106 / 20,202,090 for tiny/medium/large in both repeats;
every lane reduces allocation calls by 1,039, reallocations by 93 and
deallocations by 946. Peak above entry changes are 0 / -49,674 / -15,438 bytes
and retained-live deltas remain zero for tiny/medium/large. No allocation
increase review flag is present.

Supplementary public-API phase clocks are separate from the keep-decision
matrix and exclude setup, warmups and checks. R1/R2 p50 deltas are commit
-12.6863% / -13.0797%, transaction -3.8135% / -2.2023%, snapshot opening
+0.9812% / +0.4743%, add +3.7947% / +1.2462%, and publication -0.0322% /
+0.0312%. The proof changes only the commit validation path; the other phase
movement is diagnostic and is not attributed to the proof. Whole-process
counters include setup, warmups and checks: instructions -5.6821%, cycles
-6.4241%, branch misses -5.2313%, branches -5.3523%, cache misses -2.0587%,
context switches -6.3380% and page faults -24.1941%.

The final source review reports no semantic blocker at candidate epoch
`d18406665f4fd9ad76cb0a2530bfe8e7459bf4ada874ea353b435a84ae4f921c`, with 381
ODP tests and warning-denied all-target Clippy passing. The harness and final
gates are complete: 387 harness tests pass with one
ignored, for 768 passed ODP/harness tests in total. Warning-denied rustdoc,
scoped formatting, boundaries, precleanup and source replay, fresh-copy
portable replay, resealed +1ns tamper rejection, and owned cleanup of four
executables totaling 233,058,712 bytes all pass. The candidate adds no selector
or corpus coverage; the registry remains 439 selectors / 36 defaults, the full
non-iWork goal remains open, and iWork is excluded. See the
[comparison summary](../results/change-0463/summary.json), [phase summary](../results/change-0463/phase-summary.json),
[source review](../results/change-0463/source-review.md), and
[proof design](../results/change-0463/proof-design.md).
