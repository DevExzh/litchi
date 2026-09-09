# 0489 results review

Reusing the successful candidate XML audit inside one immutable prepared OPC
splice plan is accepted for this scoped change. Source-heavy normal owned
medians improve 20.29–21.31% and file medians improve 20.77–20.88% across the
two repeats. Authored-heavy file medians improve 14.23–14.92%; owned medians
improve 21.44–21.76%. Operation heap decreases on every measured allocator arm.
The small file-store tail regression and four RSS increases below remain open.

The [complete comparison](comparison-summary.md) retains 144 formal children,
4,320 measured samples, and all normal and allocator quantiles. The
[plot](latency-comparison.png) displays normal p50 only. Measurements are pinned
to CPU 2 and serialized under the shared lock, on an ordinary shared host.
Two process repeats give descriptive observations, not a confidence interval,
a reserved-host guarantee, or a cold-cache claim.

## Normal median latency

Each cell gives before → after milliseconds and signed percentage change.
The source/authored paragraph counts are encoded in the workload names.

| Arm | Repeat 1 | Repeat 2 |
| --- | --- | --- |
| deterministic-owned-s64-a64-short-c64 | 0.780 → 0.551 (-29.33%) | 0.788 → 0.548 (-30.39%) |
| deterministic-file-s64-a64-short-c64 | 1.064 → 0.809 (-23.97%) | 0.943 → 0.806 (-14.51%) |
| deterministic-short-read-s64-a64-short-c64 | 0.675 → 0.556 (-17.65%) | 0.676 → 0.549 (-18.82%) |
| deterministic-latency-s64-a64-short-c64 | 12.278 → 12.154 (-1.01%) | 12.269 → 12.143 (-1.03%) |
| deterministic-owned-s64-a16384-short-c64 | 98.742 → 77.252 (-21.76%) | 97.849 → 76.870 (-21.44%) |
| deterministic-file-s64-a16384-short-c64 | 155.059 → 131.920 (-14.92%) | 152.012 → 130.379 (-14.23%) |
| deterministic-short-read-s64-a16384-short-c64 | 99.432 → 77.506 (-22.05%) | 98.228 → 78.534 (-20.05%) |
| deterministic-latency-s64-a16384-short-c64 | 113.396 → 89.707 (-20.89%) | 113.314 → 89.714 (-20.83%) |
| deterministic-owned-s131072-a64-short-c64 | 386.874 → 304.451 (-21.31%) | 384.636 → 306.591 (-20.29%) |
| deterministic-file-s131072-a64-short-c64 | 386.661 → 306.354 (-20.77%) | 386.485 → 305.777 (-20.88%) |
| deterministic-short-read-s131072-a64-short-c64 | 384.027 → 306.972 (-20.06%) | 383.999 → 303.548 (-20.95%) |
| deterministic-latency-s131072-a64-short-c64 | 432.466 → 353.576 (-18.24%) | 432.931 → 352.240 (-18.64%) |
| memory_store-owned-s64-a64-short-c64 | 0.539 → 0.429 (-20.41%) | 0.538 → 0.428 (-20.34%) |
| memory_store-owned-s64-a16384-short-c64 | 68.520 → 52.198 (-23.82%) | 68.373 → 52.285 (-23.53%) |
| memory_store-owned-s131072-a64-short-c64 | 383.897 → 302.820 (-21.12%) | 383.939 → 306.036 (-20.29%) |
| file_store-owned-s64-a64-short-c64 | 3.677 → 3.491 (-5.07%) | 3.628 → 3.889 (+7.19%) |
| file_store-owned-s64-a16384-short-c64 | 76.959 → 61.181 (-20.50%) | 75.535 → 59.177 (-21.66%) |
| file_store-owned-s131072-a64-short-c64 | 386.243 → 307.429 (-20.41%) | 387.828 → 307.342 (-20.75%) |

## Adverse observations

Every elapsed-time comparison with an increase above 5% is listed below.
Both are the 64-source, 64-authored file-store arm in repeat 2; values are
milliseconds. Normal repeat 1 instead improves p50 by 5.07%. The inconsistent
repeat and the allocator tail regression do not justify dismissing this as
noise. File-store filesystem cost and host interference are plausible causes,
but this batch does not establish the cause. The regression remains a specific
follow-up; no general file-store small-input speedup is claimed.

| Binary | Repeat | p50 | p95 | p99 |
| --- | ---: | --- | --- | --- |
| normal | 2 | 3.628 → 3.889 (+7.19%) | 3.733 → 6.682 (+79.00%) | 3.880 → 7.084 (+82.60%) |
| allocator | 2 | 3.699 → 3.859 (+4.32%) | 4.254 → 5.753 (+35.24%) | 4.374 → 6.637 (+51.74%) |

Four distinct whole-child RSS observations increase above 5%. GNU time supplies
one value per process; the summary's identical p50/p95/p99 labels are not three
independent observations. Both short-read repeats increase despite lower
operation heap. These RSS results remain an unresolved tradeoff.

| Arm | Binary | Repeat | Before bytes | After bytes | Change |
| --- | --- | ---: | ---: | ---: | ---: |
| deterministic-short-read-s64-a64-short-c64 | normal | 1 | 6,447,104 | 6,950,912 | +7.81% |
| deterministic-short-read-s64-a64-short-c64 | normal | 2 | 6,639,616 | 7,041,024 | +6.05% |
| deterministic-latency-s64-a64-short-c64 | normal | 2 | 6,574,080 | 7,041,024 | +7.10% |
| memory_store-owned-s64-a64-short-c64 | allocator | 1 | 6,647,808 | 7,303,168 | +9.86% |

No other retained latency or RSS comparison exceeds +5%. No operation-heap
quantile increases. The deterministic route's median operation peak increment
falls about 19.89% (owned: 989,190 → 792,403 bytes); memory-store falls 2.14%
(9,184,134 → 8,987,347 bytes); file-store falls about 24.73%. Existing XML
workspace admission reservations remain conservatively enforced even where
actual parser allocations disappear.

Allocator counters retain all calls, bytes, live-byte conservation and failed
allocation counts. For authored-heavy deterministic owned input, repeat 1's
median allocation calls fall 131,382 → 65,826 and allocated bytes fall
12,425,667 → 7,837,665. Source-heavy owned calls fall 822 → 546 and bytes
4,069,827 → 3,659,745. These are instrumented operation measurements, separate
from whole-child RSS and normal-build latency.

## Mechanism and identity

The [implementation review](implementation-review.md) records the private
proof/target/limit capability, initial source-first and candidate audits,
preview cloning, authenticated replay, Work failure granularity, conservative
workspace admission, cancellation, exact sink counts, and final DOCX reopen.

[Whole-child profiles](profiles/profiles1-summary.md) show source-heavy
instructions falling 16.49–16.51%. Authored-heavy instructions fall 21.32% for
owned input and 18.51% for file input. These children include setup and oracle
work and are not operation-only CPU measurements. File statx calls remain
21,781 for source-heavy and 591,477 for authored-heavy input; pread64 remains
264 and 114 respectively. Owned controls retain 12 statx and 6 pread64 calls.
The reduction is consistent with removing repeated parser work, without
reducing these measured metadata or read syscall counts.

All source, authored, candidate XML, semantic and preservation oracles pass.
Candidate archive length and SHA-256 also match across phases: there are zero
candidate archive identity changes. The final source manifest is
`0f5d9a455d25146823a7a8eb3620712f6e8f199abe0d40c32678d100c3994afa`;
only the private OPC splice implementation and focused public replay tests
differ from the retained 0487 source manifest.

## Validation and retained development failure

Final OPC/DOCX all-feature tests pass with 2,002 passed and 32 ignored;
no-default-feature tests pass with 1,961 passed and 32 ignored. Format,
Clippy and rustdoc with warnings denied, crate boundaries, five serial
allocator harness tests and 14 Python helper tests pass. Five new adapter
cases bring that suite to 21 tests. The added public test covers 36 combinations
of storage mode, publication pass, and late replay length/hash failure.
ASan validation passes 27 required-positive inputs and two 10,000-run campaigns
with seeds 484 and 485 and a 65,536-byte cap.

The first development Clippy gate failed only on the grouping of a hexadecimal
test literal. Its numeric value was unchanged by the fix. The failed receipt
and earlier successful tests remain in the evidence inventory; only final
source gates are selected for the seal. The allocator harness uses serial test
execution because the prior batch retained a process-global counter race under
parallel test execution; this batch does not claim to fix that harness race.

## Cleanup and remaining work

The cleanup receipt authenticates the copied executables and evidence, archives
the generated fuzz build recipes, and removes batch-local drafts, corpus copies,
test scratch and the isolated Cargo output. Cargo's original executable path is
reused between normal and allocator builds; the retained per-role copies are
the authenticated evidence inputs. The evidence verifier must pass after
cleanup. The protected spec-gap worktree and other sessions' work remain intact.

The broader non-iWork goal remains open: CRUD/provider intersections, native
Office validation, cold-cache and concurrent workloads, remaining freshness
metadata cost, and the small file-store tail/RSS regressions require further
work. This change does not remove the standalone source audit or initial
candidate audit, and it does not establish a universal 10x improvement.
