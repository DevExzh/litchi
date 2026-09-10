# 0491: DOCX source-provider and verified-cold baseline

This batch adds reproducible evidence for the DOCX full-text read lifecycle on
owned bytes, positional file input, instrumented caller sources, capped short
reads, simulated latency/bandwidth, and filesystem cache states. It changes
benchmark code only; it does not claim a production speedup.

The measured text now survives the lifecycle clock and is checked afterward.
Document destruction remains timed; text verification/destruction is outside.
This is an explicit timing-scope revision, so historical lifecycle times are
not an equal-scope before/after baseline. A generator-only, exact-hash-gated
restoration retains the historical synthetic archive despite one known producer
namespace spelling change. Other producer drift fails closed.

The aligned verified-cold copy produces one bounded EOCD tail search. Evidence
retains its raw compressed-payload overlaps and proves exact identity, comment
padding and the sole tail probe. Open observes zero payload cache loads; main
preparation subsequently materializes one part. Prepared full-text queries stay
explicitly cold-ineligible.

[Results and individual repeat flags](../results/change-0491/results-review.md)
cover 600 provider samples and 360 filesystem samples, plus four ineligible
controls. Heap allocation and peak increments, whole-child RSS, source-call
histograms, actual short reads, source versions, timed text hashes and cold
process/residency proofs retain separate scopes. Whole-child perf/strace
profiles include corpus/setup/report work and are diagnostic only.

## ADR and correctness scope

| Constraint | Implementation/evidence |
|---|---|
| 0001, 0002, 0010, 0011, 0024 ownership | Changes remain in the separate benchmark harness; no facade/format archive dependency added. Boundary gate passes with existing iWork debt explicit. |
| 0003 immutable snapshots/source identity | Existing source-backed package APIs remain unchanged. Provider versions are compared around each operation. |
| 0005 explicit I/O/budgets/evidence | Caller-supplied providers, finite delay/bandwidth/range caps, explicit read/cache limits, serial CPU-affined captures, owned scratch and exclusive receipts. |
| 0006 preservation/security | Exact original corpus identity; bounded cold alignment proof; no normalization of production documents, cache weakening, or partial-success fallback. |
| 0008 verification | Complete harness library/allocator tests, scoped formatting, warning-denied Clippy/rustdoc, doc tests, helper tamper tests, and retained terminal gates. |

See the [reproduction commands and evidence index](../results/change-0491/README.md),
[methods](../results/change-0491/methods.md), and
[review dispositions](../results/change-0491/review-findings.md). Final cleanup
removes only this batch's compiler/scratch output and keeps authenticated final
executables. The spec-gap worktree and iWork implementation are untouched.

Remaining work includes measured range coalescing, genuine non-static borrowed
input, bounded concurrency/scaling, native producer validation, publication
intersections and broader CRUD coverage. The overall goal remains active.
