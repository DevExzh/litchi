# Independent review and disposition

Retain the final preflight implementation. The independent XLS admission
reviewer found no material correctness blocker after the reservation ordering
correction and recommends retention after reviewing final performance evidence.
Root agrees. Results are scoped observations; `performance_claim: none` means
no registered project-wide claim is promoted.

The initial geometric-reservation-then-fallback design was rejected: a failed
charge can leave partially acquired reservation-vector capacity. The final code
chooses bounded growth before one reservation attempt and abandons on failure;
allocator excess still requires charging. Successful complete scans alone can
learn intrinsic refusal, while transient pressure and failed/partial scans
remain retryable. The rejected source and checks remain archived separately.

Repeated open-plus-eight workloads improve about 75% for 54016/1 MiB,
78% for generated 70,001 cells, 31% for generated 100,001 cells and 22% for
54016/512 KiB. These gains are large relative to their paired A/A variation.
The unaffected Plan1/default drift is excluded from claimed gains.

Higher construction requests and peaks are accepted costs. Newly successful
builds retain 1,048,340 measured bytes for 54016/1 MiB and 2,096,857 for the
70,001-cell default case. Intrinsic-refusal cases retain no additional index.
Query-eight result deltas exclude the already retained owner cache. Local
logical admission is not a total-process or allocator memory bound.

Build medians regress roughly 5–8% on several routes; 54016/1 MiB build
p95/p99 regress despite improved medians. Formula-refusal phases regress about
6–9%, including non-collecting first queries whose cause is not isolated;
open-plus-eight refusal costs rise 1.80–2.44%. Tiny file-backed admission
adds a source read and regresses 1.82–2.33% over open plus eight. These costs
are explicit follow-up work, not averaged out of the result.

Corrected native-child RSS rises 7.21% for the short 512 KiB run and 4.57%
for its longer run. Other measured increases are at most 4.70%. Review accepts
these disclosed triggers; it does not infer an RSS bound or a universal memory
reduction. Initial wrapper RSS is archived and excluded.

All six final quality gates pass with 1,942 tests passed and two existing
ignored. Real/generated differential checks and separate native, allocation,
I/O, counter and source-binding audits pass. Physical cold storage, remote
latency, concurrency, cross-platform behavior and universal query improvement
remain outside this experiment. Broader GOAL work remains active.

Independent evidence review confirmed the numerical tables, examples, counts
and claim scope against the packet. Its sole documentation finding, the stale
provisional README status, was corrected before commit. Final report, coverage
and non-iWork documentation gates pass. Raw profiler reports retain their
original whitespace; edited source and documentation pass whitespace checks.
