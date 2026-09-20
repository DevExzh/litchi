# 0725 — rejected revised XLS checkpoint pilot

The revised candidate is rejected: 17 of 24 native groups fail the frozen
non-regression gates. All 16 repeated-query groups, both late-target benefit
gates, allocator checks, semantics, source-I/O checks and capture bindings pass.
The archived candidate removes work successfully, but does not qualify for
production retention. Baseline `ee5e0b0650` is restored. The non-iWork program
remains active; `performance_claim: none`.

| Owned case / metric | B1/A1 p50 delta | B2/A2 p50 delta | Mean delta range |
| --- | ---: | ---: | ---: |
| 54016 late / native q8 | −80.00% | −80.82% | −80.51% to −79.85% |
| 54016 late / repeated loop | −83.11% | −83.11% | −83.06% to −83.01% |
| 54016 missing / repeated loop | −32.15% | −32.57% | −32.67% to −32.46% |
| 54016 stored / repeated loop | −6.69% | −5.67% | −4.97% to −4.91% |
| 54016 zero-budget / native q8 | +9.08% | +9.71% | +9.07% to +9.53% |
| 54016 zero-budget / open plus eight | +8.06% | +8.22% | +8.07% to +8.29% |
| Simple first / native q2 | +5.56% | +5.56% | +5.77% to +6.64% |

Negative latency deltas are faster. Repeated statistics describe nine fresh
process loop means per leg, each measuring 50,000 queries after two preparatory
queries; they are not individual-query tails. The repeat lane uses a 2 MiB
budget, including its missing case; the native 54016-missing case uses 1 MiB.
The two comparisons are B1/A1 and B2/A2 within one A/A plus ABBA capture.

The zero-budget regressions repeat in both source modes and both pairs; no
checkpoint can be admitted on that route. Simple owned publication also fails
in both pairs. These findings suffice to reject the candidate without treating
other noisy rows as grounds to weaken the gates. Across the full matrix there
are 102 failed central-statistic checks, 205 paired tail/maximum regression
flags above 5%, nine A/A central drift flags and 21 within-phase central drift
flags. Every result remains in the [complete comparisons](results/change-0725/measurements.md).
No samples were dropped, recaptured or pooled to alter the decision. This
combined source change does not isolate the cause of scan-path regressions.

The revision builds on the rejected 0723 experiment and 0724 attribution. It
retains one target-frame chain checkpoint alongside the original worksheet and
SST checkpoints, but records the first matching raw frame in transient scan
state rather than searching collected slots afterward. After successful EOF
validation and execution/source fences, a metadata-only cursor probe reuses the
existing borrowed path. Missing indexed queries resolve their empty slot range
before allocating an unused path vector, then take the same trailing execution
and source-version checks. Duplicate slots retain their order and last-value
semantics; earlier targets use the original checkpoint and later targets may
walk a suffix. Optional checkpoint refusal preserves the original route.

Separate diagnostic traces confirm removal of 1,044, 72 and 258 FAT links on
original-target replay for 54016-late, Plan1-late and 45365-late respectively.
Construction still walks that prefix once per admitted owner. The five-query
sequence preserves exact results, source read ranges and version observations;
earlier replay retains 14 links, origin-late replay drops from 1,044 to zero,
and missing replay performs zero reads with two version observations in both
phases. Later-target behavior is covered by the integration test. All 12 primary
counted-I/O cases remain exact. This removes metadata traversal, not physical I/O.

The fixed logical charge remains 224 → 264 bytes. Typical stored-target q2 adds
40 allocated/retained bytes with no extra allocation or deallocation call and
no peak increase. Budget-capacity rounding gives smaller or negative deltas on
the 1 MiB missing and generated cases. Positive-budget missing q3/q8 remove
exactly one 16-byte allocation/deallocation pair and 16 peak bytes, with zero
retained-byte change. All other strict counters, including zero-budget and typed
refusal, retain exact parity. All 96 allocator groups pass the predeclared
40-byte q2 bound and explicit missing-warm deltas. All 34 budget-fence comparisons
preserve outcomes; Simple stored/missing budgets 489 and 490 expose changed
admission/source routes. Allocation-region peaks are not process RSS.

| Architectural requirement | Evidence / limitation |
| --- | --- |
| No decoded-value, error or source-byte cache | Private bounded metadata only; independent retained-type/source review |
| Complete validation and freshness | EOF validation, selected-frame decode, duplicate order and execution/version fences retained |
| Bounded memory and fallible reservation | Fixed charge includes checkpoint; refusal and budget fences pass |
| Cancellation | Existing checks remain; one bounded metadata walk occurs between scan and publication checks |
| Public API, dependencies, unsafe and ambient I/O | No changes |
| Concurrency | No new workers or synchronization design; no scaling claim |
| Preservation and non-iWork scope | Exact corpus results and all six repository evidence gates pass |

The baseline is `ee5e0b0650eb5b68bf0ec7cc351290d5eb69e809`. Unchanged 0684/0686
release probes run on CPU 12 with Rust 1.95.0, AMD EPYC 9R45 and warm OS caches.
Source, tool, probe, fixture and executable identities are bound at freeze and
capture completion. The matrix covers 14,400 measured fresh owners / 115,200
queries plus 432 warmup owners, 864 repeated-loop processes / 43.2 million
queries, and 576 allocator processes. Separate instrumented traces and
2-million-query profiles are diagnostic only. No cold-device, remote, native
Office, cross-platform, RSS, scaling or universal tail claim follows.

Final qualification passes formatting, checks, warning-denied Clippy/rustdoc,
1,904 CFB/XLS tests with two existing ignored tests, and 61 facade tests. All
126 real XLS fixtures in owned/file modes and all 70,001 generated cells retain
exact outcomes and source metrics. Fourteen verifier controls pass. The initial
unused-mut build failure and passing duplicate-test qualification are archived;
final builds and qualification bind the consolidated test inventory before any
main capture. Independent audit and offline replay preserve the same rejection.
Five owned build/profile roots are removed with executable identity witnesses.
[Evidence and replay index](results/change-0725/README.md).

The next bounded candidate should isolate the empty-slot setup removal on the
unchanged 224-byte index, without adding checkpoint state or scanner hooks.
Its combined-candidate benefit is promising but does not establish that the
isolated change will pass; it needs its own qualification and frozen comparison.
Further checkpoint work must account for zero-budget scan behavior before
another retention attempt. No registered claim or CRUD coverage is promoted.
