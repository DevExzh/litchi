# 0787: cached ordered Part scheduling

**Rejected and restored.** The cached path reduces latency substantially, but
one small cached case fails the frozen whole-child peak-RSS guard. No
production or test-source change is retained.

This batch tests whether completed cache hits can avoid constructing private
workers while retaining finite execution admission and ordinary prepared-read
validation. Change 0786 measured severe negative scaling on primed Parts with
zero source reads. The paired baseline for this experiment is `6b47617020`;
historical 0786 timings are context only and are not pooled into these results.

## Mechanism and scope

The candidate stays inside `litchi-opc`'s source-backed batch/cache owner.
After the existing preparation, output reservation, scheduler reservation,
CPU-task charge, and parallel-entry fence, it checks whether all requested
cache entries are complete. The query excludes active flights and provisional
publication, takes no payload clones, and changes no cache statistics or LRU
state. All-ready private batches execute ordinary prepared reads in logical
wave order. Every request retains its source/context checks and panic
translation; a failed wave drains before the lowest ordinal error is returned.
Later waves stop. A stale hint can become an ordinary cold read.

Caller-supplied scoped workers retain their callback route. Budget admission
and conservative scheduler reservations remain in place even when threads
are avoided. Thread/channel construction and their internal allocation or
thread-start failures disappear on the all-ready private path. No dependency,
public API, global executor, cache authority, or retention policy is added.

The matrix has 60 cases: 32 Parts of 4 KiB, 256 KiB, or 31 large plus one
small member; fresh and primed sessions; task floors 0 and 64 KiB; requested
widths 1/2/4/8/32. Inputs are immutable memory. Fresh does not mean physical
cold. The wall interval is the ordered read, excluding corpus construction,
metadata setup, priming, verification, and returned-batch drop. CPU time
surrounds that interval; peak RSS covers the whole child process.

Six native paired blocks use fixed before/after orders, 30 samples, and three
warmups per child. Two separate source-observer blocks use two samples each.
Sixty qualification children run on each leg. The frozen rule rejects any
p50 or RSS case whose median paired after/before ratio exceeds 1.05 and whose
bootstrap lower endpoint exceeds 1; benefit requires an eligible cached case
with at least 3% median improvement and an upper endpoint below 1. Ten thousand
bootstrap resamples use seed 787078. Tails remain individually visible.

## Accepted ADR compliance

All 32 accepted ADRs, their index, the goal and scenario taxonomy retain the
35 input hashes in the packet. The scoped compliance assessment is:

| ADRs | Assessment |
| --- | --- |
| 0001, 0002, 0024 | Low-level OPC owner only; layering and dependency direction unchanged. |
| 0003 | Immutable reads and returned payload ownership unchanged; no edit, patch, commit or merge behavior. |
| 0004, 0025–0027 | No public semantic/editor API change; typed errors retained. |
| 0005 | Existing source cache remains authoritative; bounded explicit execution and measured adoption gate. |
| 0006, 0008 | Ordinary validation and source fences retained; focused tests, exact bytes/order, and applicable crate gates required. |
| 0007, 0009, 0012–0023, 0028–0030 | No semantic model, ODF, iWork, rendering, facade, format-writing, or unrelated architectural boundary changes. |
| 0010, 0011 | OPC uses its existing archive/source abstractions; ZIP validation and output ownership unchanged. |
| 0031 | Worker/I/O reservation, CPU-task charges, task-size floor, caller facility and release semantics retained. |
| 0032 | No new derived memo or payload retention; read-only cache scheduling hint adds no second authority. |

This is low-level CRUD category 15 evidence. It does not promote native-format
CRUD coverage or close physical-cold, delayed-range, cross-session contention,
allocation, hardware-counter or whole-program requirements. OLE2/OOXML remain
active, ODF remains deferred by the recorded owner decision, and iWork is
excluded.

The [packet](results/change-0787/README.md) contains the frozen plan, exact
source archives, raw measurements, replay tools, and limitations.

## Baseline mechanism evidence

Six operation-wrapper Callgrind regions qualify: width-one Ir is 27,877 in
both repeats, width-eight Ir is 58,890/57,924, and width-32 Ir is
96,902/98,223. The dominant caller paths at widths 8 and 32 descend through
private thread creation and TLS setup. At width 32, `pthread_create` and
allocator work are prominent self-cost rows. These are guest instruction
diagnostics, not native CPU shares; worker roots can be disconnected and
owner-inclusive cost is not an all-thread total.

The first Callgrind attempt crashed before the owner in rustix CPU-clock
initialization and collected zero Ir. Its raw failure, original probe and
binary identity are retained. A profile-only compatibility probe returns no
CPU-clock value; the native benchmark is unchanged. Both successful regions
and the failed attempt have independent offline custody checks.

Whole-child syscall traces corroborate private-thread creation. Before the
change, fresh width-8/32 children create 8/32 threads, while primed children
create 16/64 (preload plus measured operation). Width-one children create
none. After the candidate, primed width-8/32 children create only 8/32
threads for preload; fresh counts remain 8/32. Both repeats agree, and
source reads remain 64 for fresh operations and zero for primed operations.
These diagnostic timings are excluded from paired native results. The
[profile analysis](results/change-0787/profile-analysis.md) gives the raw
counts, self-cost partition, graph limitations and source-count checks.

## Correctness evidence

The candidate adds seven focused tests: completed/flight/pending cache-state
hints without LRU/hit mutation; deterministic eviction after a ready hint
followed by exact-byte cold fallback; cached single/multiwave execution with
zero reads and caller-thread version callbacks; explicit caller-facility
callbacks; multiwave panic translation/drain/stop; one-wave double-panic
lowest-ordinal selection; and a source change after the helper entry fence.

The first compile preflight found two denied explicit Arc casts in new tests.
The corrected candidate uses ordinary coercions. The failed preflight log,
source manifest and complete changed source files remain retained. Final OPC
all-feature testing passes 994 tests with one ignored across 40 summaries,
including all seven additions. All six quality gates pass: formatting, all-feature/all-target checking,
all-feature tests, warning-denied library Clippy, warning-denied rustdoc, and
the repository crate-boundary checker. Exact command receipts accompany the
compiled candidate source.

## Paired results and decision

All 1,080 reports / 22,200 samples pass exact output/hash/order, source-work,
CPU-task and permit-release checks. The independent raw audit reconstructs
96 logical payloads and reproduces all 60 paired p50/p95/p99/RSS estimates
and confidence intervals. Sixteen eligible cached cases reduce median paired
p50 by 96.75–99.03%. No p50 case exceeds the 5% median regression threshold.

The large, zero-floor cache control illustrates the mechanism. Absolute
columns are medians of six process p50s; percentage changes use paired block
ratios, so dividing the absolute columns need not reproduce the percentage.

| Requested width | Before p50 (µs) | Candidate p50 (µs) | Paired p50 change | After/before 95% CI |
| ---: | ---: | ---: | ---: | --- |
| 1 | 3.025 | 2.975 | -1.918% | [0.930354, 1.047592] |
| 2 | 228.566 | 6.955 | -96.968% | [0.028047, 0.033904] |
| 4 | 232.631 | 7.460 | -96.774% | [0.030463, 0.033258] |
| 8 | 296.796 | 8.135 | -97.246% | [0.026185, 0.029323] |
| 32 | 589.352 | 7.260 | -98.754% | [0.011858, 0.013058] |

The candidate nevertheless fails the frozen RSS rule for **small / primed /
floor 0 / width 4**: median paired ratio **1.056969**, bootstrap 95% interval
**[1.032110, 1.074752]**. All six child pairs increase:

| Block | Before peak RSS (KiB) | Candidate peak RSS (KiB) |
| ---: | ---: | ---: |
| 0 | 4108 | 4276 |
| 1 | 4116 | 4400 |
| 2 | 4084 | 4276 |
| 3 | 4124 | 4456 |
| 4 | 4124 | 4400 |
| 5 | 4116 | 4212 |

The corresponding width-32 small cache case also has a high median RSS ratio
(1.082901), but its interval [0.971751, 1.139141]
crosses one and therefore does not independently reject under the frozen
conjunctive rule. Both observations remain visible.

Peak RSS includes code pages, corpus construction, priming, allocator state,
verification and teardown. This packet does not isolate the cause of the
small-process increase or measure operation allocation. The result is not
dismissed as noise, and the large latency benefit does not override the
prespecified memory guard. The original production and test files were
restored; all 9,196 production-file hashes match the baseline.

Three paired p95 medians and seven p99 medians exceed a 5% regression. The
largest p95 ratio is 1.2232 (mixed, primed, floor 0, width 1); the largest p99
ratio is 1.3603 (large, primed, floor 0, width 1). These short-duration serial
controls and every individual block spread remain in the
[60-case table](results/change-0787/paired.md) and
[CSV](results/change-0787/paired.csv). No general tail-latency improvement or
whole-program speedup is claimed.

The first full offline replay incorrectly required exact equality of observed
peak simultaneous reads on fresh parallel operations. That scheduler-dependent
observation is now retained separately for each leg and bounded by the worker
limit; deterministic source work, zero active reads after return, output and
resource checks remain exact. The independent audit already passed and found
the same RSS rejection before this checker correction. No capture or adoption
threshold was changed.

## Remaining work

The thread-creation hypothesis is supported by the source, scoped diagnostics,
thread census and cache-hit timings. A follow-up must attribute the small-case
whole-child memory increase and meet a freshly frozen representative memory
policy before this implementation can be reconsidered. Additional public
Office CRUD, physical-cold/range-source and cross-session measurements remain
necessary; this rejected low-level candidate closes none of those gaps.

## Integration and cleanup

The original production tree is restored and the rejected candidate remains
archive-only. Offline paired validation and profile/trace replay pass after
target removal. All six executable identities were verified before deleting
1,285,185,572 bytes of owned target files; `cleanup.json` preserves their
custody. Only this batch's worktree, branch, copied lockfile, links and bytecode
caches are eligible for cleanup. The three unrelated main-worktree files and
all pre-existing worktrees remain outside this batch. Final commit and
relocation checks are recorded after integration.
