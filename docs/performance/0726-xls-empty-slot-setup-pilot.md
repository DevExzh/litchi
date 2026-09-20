# 0726 — rejected XLS empty-slot setup pilot

The isolated candidate is rejected by four native mean-latency checks across
four of 24 groups. Every paired native p50 check and all 16 repeated-query
groups pass. Both missing-query benefit gates, all allocator, semantic,
source-I/O, budget-fence and binding gates pass. Production is restored to
`3ad29e42da`; the exact candidate and every observation are archived.
`performance_claim: none`. The non-iWork performance program remains active.

| Owned missing-query measurement | B1/A1 p50 delta | B2/A2 p50 delta | Mean delta range |
| --- | ---: | ---: | ---: |
| 54016 native q8, 1 MiB index budget | −25.00% | −25.00% | −27.68% to −26.03% |
| 54016 repeated loop, 2 MiB index budget | −32.43% | −33.12% | −33.09% to −32.51% |
| Simple native q8, 2 MiB index budget | −25.00% | −25.00% | −23.83% to −23.44% |

Negative deltas are faster. Repeat statistics summarize nine fresh process loop
means per leg, each measuring 50,000 queries after two preparatory queries;
they are not individual-query percentiles. File-backed 54016 missing loops
improve 5.17–5.66% in p50 and about 5.27% in mean. Owned 54016 missing
open-plus-eight p50 improves 8.52–9.17%, but this combined workflow includes
unchanged opening and scanning; its change cannot be attributed solely to the
small allocation removal.

| Failed native mean check | Pair | Baseline → candidate mean (ns) | Regression |
| --- | --- | ---: | ---: |
| 54016 stored, owned q3 | B2/A2 | 1,034.51 → 1,124.30 | +8.68% |
| 54016 stored, file q3 | B1/A1 | 2,578.82 → 2,722.90 | +5.59% |
| Generated 70,000, owned q3-to-q8 mean | B2/A2 | 554.25 → 586.82 | +5.88% |
| 45365 late, file q8 | B2/A2 | 1,719.20 → 1,863.60 | +8.40% |

Each exceeds both the 5% percentage limit and the allowed 10 ns warm-query
exception. These four failures still reject retention; passing medians do not
replace a required mean gate. There are also 119 paired p95/p99/maximum flags
over 5%, 14 A/A central drift flags and 28 within-phase central drift flags.
All samples and [complete comparisons](results/change-0726/measurements.md)
remain available. No sample was removed, replaced or reweighted. The evidence
does not establish the cause of the failed means or justify calling them noise.
The [independent distribution review](results/change-0726/distribution-review.md)
finds candidate maxima of 7,820, 11,230, 2,346.67 and 12,750 ns in the four
failed comparisons, versus baseline maxima of 1,590, 2,820, 865 and 1,800 ns.
Mixed within-phase changes leave attribution unresolved. A targeted replication
would be diagnostic only; future retention still requires full-matrix gates.

This candidate changes only `replay_indexed_cell`. It resolves the existing
borrowed slot range after worksheet lookup. For an empty range, it performs
the same trailing execution and source-version checks and returns `None`
before allocating an unused workbook-path vector or preparing unused resolver
and chain hints. The non-empty loop uses the same slots in the same order and
retains selected-frame decoding, duplicate/error precedence and all fences.
The scanner, index admission, 224-byte fixed charge, cache layout and every
other Rust file remain unchanged. This isolates the missing-query setup seam
from the rejected checkpoint/scanner changes in 0725.

All eight positive-budget missing q3/q8 allocation rows remove exactly one
16-byte allocation and deallocation, and 16 bytes from the allocation-region
peak, with no retained-byte change. Every other allocator counter, including
q2, zero-budget, stored and typed-refusal routes, is exactly equal. All 96
groups pass. All 12 primary counted-I/O cases and all 34 budget-fence observations
preserve outcomes and exact source metrics; unlike the checkpoint experiments,
there is no admission-route exception at budgets 489 or 490. Region peaks are
not process RSS measurements.

| Architectural requirement | Evidence / limitation |
| --- | --- |
| Cache semantics and complete validation | Published index still follows a complete successful scan; no scanner or publication changes |
| Freshness and cancellation | Entry, per-slot and final execution checks remain; missing branch retains final source check |
| Error and duplicate ordering | Worksheet lookup precedes the early return; stored replay is otherwise byte-for-byte equivalent |
| Memory and cache ownership | No new retained state; unchanged 224-byte charge; exact budget and allocator gates |
| API, dependencies, unsafe, ambient I/O and concurrency | No changes; no new scaling claim |
| Preservation and scope | Exact corpus parity and all six non-iWork repository evidence gates pass |

The source and independent audits explicitly verify fence ordering. Existing
`an_indexed_missing_target_still_takes_the_trailing_freshness_fence` injects
mutation at the second version observation and verifies `SourceChanged` with
zero reads. Warm/cold and concurrent-handle tests cover values and missing
results. No timing-dependent cancellation test was added.

Baseline `3ad29e42dade8cf51c09b498f4085066683a902b` and candidate use unchanged
0684/0686 release probes, Rust 1.95.0, CPU 12 on AMD EPYC 9R45 and warm OS
caches. Source, tool, fixture, probe and executable identities are frozen and
rechecked at capture completion. The full matrix covers 14,400 measured fresh
owners / 115,200 queries plus 432 warmup owners, 864 repeated-loop processes /
43.2 million queries, and 576 allocator processes. Separate 2-million-query
missing-result perf profiles are diagnostic only. The local deletion needs no
new instrumented checkpoint trace; direct source review and exact allocation
and source metrics establish the work removed. No cold-device, remote, native
Office, cross-platform, RSS, concurrency or universal tail claim follows.

Formatting, all-target/all-feature checks, warning-denied Clippy and rustdoc
pass. CFB/XLS tests total 1,900 passed and two existing ignored; facade tests
add 61 passed. All 126 real fixtures in owned/file modes and all 70,001 generated
cells retain exact corpus results and source metrics. Seventeen verifier
controls pass. Independent audit and post-cleanup replay reproduce rejection.
Three owned build/profile roots are removed with exact executable witnesses.
[Evidence and replay index](results/change-0726/README.md).

The next step is distribution-level attribution of the four failed warm means
and an independently specified replication design before another retention
attempt. The present rejection remains authoritative. The observed allocation
removal and missing-query benefit alone do not establish whole-matrix
non-regression. No registered claim or CRUD coverage is promoted.
