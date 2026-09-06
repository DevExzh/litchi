# 0436: bounded ordinary-text Work spans in ODT

Retain the private ODT text batching change. In the named fresh paragraph
streaming scenario, normal p50 is descriptively 19.382–20.055% lower for 64
paragraphs, 29.886–30.156% lower for 8,192, and 29.602–29.896% lower for
32,768 across two repeats. Every archive/content/styles/meta/semantic/sink
identity matches. Allocator vectors are identical after alignment by sample
index; operation peak above entry remains 420,091 bytes at every size.

## Mechanism and correctness

`ParagraphWriter` previously consumed hierarchical Work and wrote scratch XML
for each ordinary scalar. It now borrows ordinary UTF-8 spans of at most 256
bytes and charges their exact byte length once. Spaces, controls, entities and
tags retain their existing encoding. Rejected checked arithmetic, paragraph
or composed-content bounds, and Work charges fall back to the scalar writer,
preserving the first failing scalar and content-first refusal precedence.
The budget owner rolls back a rejected cumulative charge through ancestors.

Cancellation is polled for every scalar admitted to a span. An already charged
span can finish copying up to 256 bytes before the next check; internal scratch
progress is not promised to match a scalar asynchronous timeline. The failed
fragment/publication contract still reports only accepted caller sink bytes.
No public API, dependencies, allocation strategy, unsafe code, hidden workers,
common archive ownership or validation policy changed.

## Reproducible experiment

[The bundle](../results/change-0436/README.md) retains 24 reports, 720 samples,
four formal large normal profiles, and six pilots per build. CPU is AMD EPYC
9R45, Linux 7.0.0-1011-aws, Rust 1.98.1, CPU affinity 2, one worker. Builds use
`-Cforce-frame-pointers=yes`, release debug info, four Cargo jobs and disabled
incremental compilation. Normal and System-allocator-instrumented binaries
are distinct. Workloads use warm in-memory deterministic mixed Unicode/entity
paragraphs, a 4,096-byte authoring window and HashingDiscardSink.

Before build revision: `4eb78e97ddafa97bf527383d7a59c4b0249d279d`.
Candidate implementation: `524db2696efc77d85f0abdfd5c21f39a6d5ec18a`.
The before build reused cached binaries exactly matching 0435 after; its
45.74% consume-self profile is retained with original provenance as preparation.
All formal captures use candidate ambient checkout state, with executable
sources independently bound in each build descriptor. Each report retains
30 samples after three warmups; order is A1/B1/B2/A2. The immutable copied
0435 oracle validates both roles as `after-streaming`, and the outer verifier
requires exact cross-revision identities.

The operation timer includes fresh paragraph construction, authoring,
publication and sink writes; it excludes corpus setup, reopen/semantic gates,
sink finalization and digest extraction. Profiles and GNU time RSS include
the whole process. The build/capture/check commands and hashes are retained in
receipts; [design](../results/change-0436/design.md) predates implementation.

## Individual normal results

All times below are milliseconds. Full means, throughput, 95% Student-t mean
intervals and fixed-seed bootstrap mean/median intervals remain in
[summary.json](../results/change-0436/summary.json). The six p50 bootstrap
interval pairs do not overlap, but they resample observations within one
process per report; two repeats do not justify a broad population claim.

| Paragraphs | Repeat | Before p50 | After p50 | Change | p95 before → after | p99 before → after |
| ---: | :--- | ---: | ---: | ---: | ---: | ---: |
| 64 | R1 | 0.168176 | 0.135580 | -19.382% | 0.181951 → 0.143781 | 0.183021 → 0.153941 |
| 64 | R2 | 0.170406 | 0.136231 | -20.055% | 0.179671 → 0.144411 | 0.181391 → 0.145491 |
| 8,192 | R1 | 13.737551 | 9.594912 | -30.156% | 13.800611 → 9.636741 | 13.857182 → 9.658062 |
| 8,192 | R2 | 13.649905 | 9.570543 | -29.886% | 13.709810 → 9.631113 | 13.718340 → 9.652822 |
| 32,768 | R1 | 54.250537 | 38.191235 | -29.602% | 54.759954 → 39.131889 | 55.072056 → 39.213499 |
| 32,768 | R2 | 54.693847 | 38.342484 | -29.896% | 54.921722 → 38.574520 | 54.968873 → 38.588860 |

Paragraph throughput p50 rises 24.041–25.086% for tiny, 42.624–43.175% for
medium and 42.050–42.646% for large. These are descriptive observations for
this scenario and environment. The earlier Builder/streaming API tradeoff is
not remeasured here, so no current matched Builder comparison is claimed.

## Allocation, RSS and regressions

Every aligned allocator vector matches before/after, including allocation and
deallocation traffic, entry/exit live values and region peak. Allocation calls
are 577 / 51,377 / 204,977 and requested bytes are 1,692,442 / 3,041,690 /
7,121,306 for tiny/medium/large. Region peak above entry is exactly 420,091
bytes in every retained allocator sample. Live delta and balance are zero.

GNU time whole-process RSS ranges from 84,537,344 to 84,729,856 bytes and is
essentially unchanged. Reported process high-water vectors remain a separate
scope. No operation-peak-to-RSS inference or RSS improvement is claimed.

None of 78 matched comparisons crosses the 5% regression threshold. One of
78 repeat comparisons is flagged: candidate tiny normal p99 changes from
153,941 to 145,491 ns (−5.489%). This is downward repeat drift, retained as a
noise review flag. No result is hidden in a geometric mean. The independently
rederived [decision](../results/change-0436/decision.json) retains the change
based on useful repeated latency reductions and unchanged byte/resource gates.

## Whole-process diagnostic profiles

Before → after perf stat: instructions 35,933,377,833 → 24,602,394,259;
cycles 9,725,019,797 → 7,271,502,179; branches 6,742,726,422 → 4,606,931,198;
branch misses 18,605,575 → 8,919,139. IPC is 3.695 → 3.383 and branch-miss
rate 0.276% → 0.194%. All retained running percentages are 100%.

The fresh formal record shows `ExecutionContext::consume` self share
44.27% → 19.24%. Candidate SHA hashing is 13.70%, XML audit 8.28%, memset
8.16%, fragment-shape validation 6.31%, deflate 4.77%, CRC 4.35%, context
check 3.87%, paragraph check 3.28% and paragraph emission 2.15%. This supports
the proposed reduced-accounting mechanism, but whole-process samples include
setup and oracle work and are not an operation-only causal or Amdahl fraction.
Both record lanes report zero lost samples, with 7 / 9 addr2line warnings;
those warnings concern symbolization. L1 misses were reported as zero without
a load denominator; LLC was not collected. Neither supports zero-miss claims.

## Validation and remaining scope

All 1,002 final ODT release tests pass, none ignored, including the new scalar
chunk oracle and exhaustive representative limit sweeps. Minimal-feature
check, scoped all-feature/all-target Clippy, warning-denied docs, owned
formatting and crate-boundary gates pass. The pre-existing large-enum lint
allowance is explicit. The first test compile's eight unnecessary-path errors
remain as a failed receipt and were corrected before the passing runs.
[Source review](../results/change-0436/review.md) found no release blocker.

The unchanged common/harness suites retain their prior 0435 full validation;
the harness was rebuilt for this provider and all current report oracles pass.
Native fixtures are functional coverage, not a fresh native Office execution.
The capture-time verifier is retained exactly under `versions/`: replay fixes
an inherited five-directory cleanup expectation to the actual three paths
without rewriting any capture receipt or measurement.

Copied verification passed before and after cleanup, including rejection and
restoration of all eight mutation probes. Exactly three temporary directories
were removed (1,822,597,957 regular-file bytes), preserving both shared build
directories and the pinned goal. Portable replay requires no retained binary.

This closes the immediate measured ODT scalar-accounting follow-up. ODP fresh
creation is the next coverage target; richer authoring, existing-document append,
Part addition, repackaging, native breadth, cold/range I/O, scalable parallelism
and the other original goal criteria remain open. No change to docs/GOAL.md,
iWork scope, or the original completion requirements is made.
