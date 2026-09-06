# Review of latency, work and retention

All 23 primary >5% absolute paired/repeat changes and all 10 confirmation flags
have individually bound dispositions in `regression-review.json`. Both primary
range/media API medians improve 31.324%/31.048%, exceeding the frozen 10% gate.
Plain API medians change less than 0.4%; no primary RSS flag exceeds 5%.

Primary bytes/media R1 baseline is 40.866 ms versus 28.326 ms in R2. Its cause
is unresolved; the 36.419% R1 improvement is not a stable CPU claim. The separate,
post-primary frozen ABBA investigation retains four fresh processes and 120
samples. It confirms about 8% API improvement and about 6% planning overhead.
Planning now performs compressed capture/verification; publication removes that
work. All original samples remain visible, with no pooled estimate or replaced
primary acceptance gate. p95/p99 improvements are conditional on these processes.

Source publication changes from 425 data-read calls / 16,786,581 returned bytes
to zero. Source planning changes from 560 / 16,788,178 to 441 / 16,791,887; total
source reads including open change from 1068 / 33,584,577 to 524 / 16,801,705.
Destination reads remain identical. Publication still has 23 source-cache hits
and zero cold loads; the cache is 64 MiB / 128 entries, not a bypass experiment.
Source work decreases from 50,366,359 to 33,589,143 units after open.

Source memory reservations at planned and published checkpoints increase from
16,807,458 to 33,622,602 bytes: an explicit 16,815,144-byte retention cost until
the plan drops. Both then return to 16,807,458. Existing staged decoded copies
remain beside cache and compressed bytes. Whole-process RSS is dominated by
untimed fixture construction and does not establish a timed allocator peak or
prove no regression for oversized/cache-bypass/tight-budget cases. Capture may
charge compressed work before decoded cache-memory admission fails; memory-only
fallback remains bounded but is not an exhaustive tight-work-budget equivalence
claim. Shared staged decoded ownership and admission order are follow-up work.

The public PPTX editor is consuming and lineage-bound. Reusable OPC captures
support the borrowed plan rerun; this does not add public repeat-publication
support. Native application breadth, cold I/O, scaling and broader CRUD remain
outside this batch. The full non-iWork goal remains active.
