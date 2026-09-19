# 0686 — bounded XLS index growth and retry admission

Status: final validation passed; scoped performance disposition below. `performance_claim: none`.
OLE2/OOXML remain active; iWork is excluded.

## Hypothesis and baseline

The bounded occurrence cache from [0684](0684-xls-occurrence-query-index.md)
retries optional construction on every selected query after the first miss.
Retries are necessary after transient budget pressure or cancellation, but a
worksheet whose minimum locator storage exceeds the owner's immutable local
ceiling cannot fit on a later call. [0685](0685-xls-worksheet-chain-checkpoints.md)
identified below-capacity retries as remaining work.

The 54016 worksheet contains 38,950 occurrences: its intrinsic storage fits
1 MiB, but doubling capacity to 65,536 slots does not. That case requires a
bounded growth step, not permanent suppression.

A fresh baseline at `fb0416104` confirms the cost on 54016 with a 1 MiB local
ceiling: the eighth query still requests 1,703,330 allocation bytes over 28
calls, with 851,992 peak live bytes and only four retained result bytes.
Disabling the cache instead requests 130,658 bytes over 15 calls on that query.
These are separate allocator captures, not native latency estimates.

## Bounded growth

Before reserving slot growth, the candidate chooses the smaller of the geometric
step and the remaining owner-local space after fixed cache metadata, current
candidate reservations, and the next reservation-vector growth.
The normal local and hierarchical reservation checks still run. Resident and
in-flight competitors are not ignored by admission; the computed size is only
an upper bound for one attempt. Allocator excess capacity remains charged
before publication. Any failed reservation abandons the candidate. Review
rejected an initial retry-after-failure implementation because a failed charge
can leave partially acquired reservation-vector capacity; that state must not
be reused. Its source and passing normal-allocator checks are archived, without
a performance claim.

This allows the 1 MiB 54016 case to retain an index if its actual allocation
fits, instead of falsely treating a doubling failure as a permanent size fact.
The measured retained cost of that newly admitted index must be reported.

## Completed-scan admission proof

Only the index-collecting scan counts observed cell occurrences. It continues
counting after optional storage is abandoned, without retaining additional
slots. A complete successful scan and final source/execution fences precede
learning. The intrinsic lower bound is the fixed cache metadata plus index
metadata plus occurrence count times logical slot weight; duplicate coordinates
remain separate occurrences. Checked arithmetic preserves the bound.

If that minimum exceeds the fixed owner-local ceiling, a sentinel in the
existing fixed worksheet hotness table suppresses future optional collection.
Those queries still run the ordinary complete validated scan. This stores no
value, error, source bytes or additional table. No partial scan learns a refusal.
The immutable source owner supplies identity; this is not a cross-owner cache.

Temporary execution-budget pressure, cancellation, pinned or concurrent entries,
allocation failures, reservation-vector storage and spare slot capacity cannot
by themselves establish this intrinsic bound. They remain retryable. Limits
and source versions are not relaxed; unsuccessful queries are not memoized.

## Measured result

The [packet](results/change-0686/README.md) binds baseline `fb0416104`, both
owner sources, probes, fixtures, binaries and final checks. Native A/A and
A/B/B/A use 100 fresh owners per leg after three warmups: 26 groups, 15,600
owner records and 124,800 query records. Timers cover opening plus individual
queries; source construction, benchmark outcome projection and serialization
are outside them. File input has warm OS caches.

Two deterministic row-major fixtures exercise the default 2 MiB limit:
70,001 occurrences fit with bounded growth; 100,001 intrinsically exceed it.
Two generation runs produce identical bytes. Independent full visitors verify
every stored cell through counts and semantic digests for owned/file sources.
These are scanner fixtures, not native Office compatibility evidence.

The table reports the sum of opening and eight query timers. Percentages are
the two paired median comparisons; all individual phases, means, tails, A/A
spreads and median bootstrap intervals remain in the packet.

| Workload and local limit | Owned | File |
|---|---:|---:|
| 54016 stored, 1 MiB | −74.86% / −74.86% | −74.50% / −74.46% |
| 54016 missing, 1 MiB | −75.01% / −74.98% | −74.74% / −74.71% |
| 54016 stored, 512 KiB | −22.00% / −22.02% | −21.83% / −22.20% |
| Plan1 stored, 32 KiB | −19.62% / −19.37% | −16.75% / −16.75% |
| Generated 70,001 cells, default | −78.24% / −78.23% | −78.02% / −78.01% |
| Generated 100,001 cells, default | −31.01% / −30.98% | −30.70% / −30.88% |
| 54016 stored, default control | −0.06% / +0.56% | +0.28% / +0.70% |
| Simple stored, default control | −2.57% / −3.31% | +0.62% / +0.68% |

For example, owned 54016/1 MiB falls 4.677/4.650 → 1.176/1.169 ms;
its eighth query falls 566.6/563.5 → 1.480/1.475 µs. Generated 70,001-cell
open-plus-eight falls about 8.11 → 1.76 ms. The 100,001-cell case continues
scanning, but avoids collecting and abandoning an index on every later call;
its eighth query falls about 1.45 → 0.862 ms. This is useful both when an index
can fit and when a validated lower bound proves it cannot.

## Costs and regressions

[All paired median triggers](results/change-0686/regressions.md) and
[additional mean/tail triggers](results/change-0686/tail-regressions.md) remain
visible. The second query now fills more available capacity: it regresses
7.24–8.38% for 54016/512 KiB, 5.51–5.75% for owned Plan1/32 KiB,
6.71–7.25% for the 70,001-cell fixture and 7.03–7.46% for 100,001 cells.
The 54016/1 MiB construction median improves, but its stored/missing p95/p99
regress about 7–10%; no uniform construction or tail-latency gain is claimed.

Formula-refusal phases regress roughly 6–9%, including first queries that do
not collect an index. Their open-plus-eight windows regress 1.80–2.44%.
The cause of the first-query shift is not isolated; it is not attributed to
index reuse. Errors remain uncached and subsequent calls still validate.
Simple/450-byte file input also gains one source read on indexed hits and
regresses 1.82–2.33% over open plus eight queries, despite its owned-source gain.

The unaffected owned Plan1/default A2 control drifts 20.36% over open plus
eight; its apparent −16.45% second-pair improvement is not a claimed gain.
Bootstrap intervals describe sampling within each leg and do not eliminate
shared-host drift. With 100 samples, p99 remains a descriptive sample tail.

## Allocation, retention, I/O and CPU work

All 104 allocation groups agree across three repeats per binary. Successful
admission intentionally retains memory that baseline failed builds released.
At the build query:

| Owned build | Requested bytes before → after | Peak live bytes before → after | Retained bytes before → after |
|---|---:|---:|---:|
| 54016 stored, 1 MiB | 1,703,330 → 2,751,666 | 851,992 → 1,113,784 | 4 → 1,048,340 |
| Generated 70,001, default | 3,341,794 → 5,438,642 | 1,638,438 → 2,162,310 | 9 → 2,096,857 |
| Generated 100,001, default | 3,341,794 → 5,438,530 | 1,638,438 → 2,162,310 | 9 → 9 |

The extra construction requests/peaks are real costs. At query eight, 54016/1 MiB
requests 1,703,330 → 102 bytes (28 → 8 calls), and generated 70,001 requests
3,341,794 → 122 (30 → 8 calls). Generated 100,001 requests
3,341,794 → 196,258 (30 → 16 calls), with peak live bytes
1,638,438 → 65,574. Query-eight retained deltas measure the result only; the
previously built index remains owned outside that gauge window. No new cache
table is allocated and existing fixed/index logical charges remain 128/192.

Counted opening and first/build-query source operations match exactly on all
13 cases. Newly admitted indexed hits use the existing replay path: 54016/1 MiB
query eight falls from 26 reads/615,829 bytes to 3 reads/21 bytes. The generated
70,001-cell hit falls from 27 reads/1,260,323 bytes to 3 reads/26 bytes.
The 512 KiB and 100,001-cell cases preserve their full-scan reads and freshness
checks exactly. Existing source checks are unchanged; fewer indexed reads also
mean fewer physical-read freshness checks, without removing operation fences.

Separate whole-process counter runs compare 10 and 1,010 queries, repeated
three times, with one warmup owner and one measured owner. Subtraction is divided
by 2,000 extra queries and includes semantic projection/reporting; these are
not native API timings. Extra-query instructions fall 99.59% for 54016/1 MiB,
99.87% for generated 70,001, 15.16% for 54016/512 KiB and 13.72% for generated
100,001. The tiny default control changes −0.06% in instructions.

Corrected native-process peak RSS medians (10/1,010 queries, KiB) are:

| Case | Before | After |
|---|---:|---:|
| 54016, 1 MiB | 5,132 / 5,300 | 5,224 / 5,420 |
| 54016, 512 KiB | 4,436 / 4,816 | 4,756 / 5,036 |
| Simple, default | 2,536 / 2,796 | 2,520 / 2,760 |
| Generated 70,001 | 5,608 / 6,000 | 5,608 / 5,852 |
| Generated 100,001 | 6,296 / 6,540 | 6,592 / 6,744 |

The short 512 KiB run is a +7.21% RSS review trigger (+4.57% at the longer
length). Allocator gauges, logical weights and process RSS are different
quantities; these results do not establish an RSS bound. Initial captures timed
the enclosing perf process and measured its roughly 62 MiB footprint. They are
archived and excluded: final commands put `time` inside `perf` so RSS covers
the probe child. Counter measurements still include the small timing wrapper.

The baseline profile attributes 23.69% self samples to candidate slot collection.
Equal 20,000-query profiles retain 11,414 baseline versus only 75 candidate
samples after admission; the candidate is too short for precise function-share
claims. User-space reports are retained; restricted kernel symbols limit kernel
attribution. Hardware counters and phase timings carry the quantitative result.

## Validation and disposition

All six final quality gates pass: workspace formatting, owner/all-target check,
warning-denied Clippy, 1,881 CFB/XLS tests (two existing ignored), 61 facade
tests and warning-denied rustdoc. The seven new unit and four integration tests
cover bounded growth, completed-count learning, dropped/failed/cancelled scans,
managed and pinned pressure, overflow, source changes and target formula errors.
The 126-real-fixture differential matches for both sources; generated full
visitor counts and semantic digests also match. Dependency, strict/structural
claims, report classification, coverage and non-iWork gates pass.

Independent source review approved the corrected preflight ordering. Independent
performance review recommends retaining the change with the disclosed build,
retention, refusal and RSS costs; see the [review](results/change-0686/review.md). Physical cold-cache,
remote latency, parallel throughput and cross-platform behavior remain outside
this experiment. SST/within-sheet traversal, refusal overhead and broad program
completion remain open. No registered claim or coverage row is promoted.
