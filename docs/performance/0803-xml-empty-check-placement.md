# 0803 — exact-empty check placement diagnostic

This experiment compares the rejected 0802 helper with the same implementation
after moving its exact-empty decision from construction to the first iterator
request. Neither leg is current production. The result cannot qualify a
production candidate or authorize workflow trials; 0802 remains rejected.

The constructor now always stores `First`. At the start of `next_first`, the
same raw-tail length equality sets `Done` and returns `None` for exact-empty
tails. Whitespace and malformed tails still reach the parser. The short prefix,
bounded linear array, ordered map, key preflight, error positions and clone
implementation are unchanged. Both iterator layouts measure 128 bytes; this
is object size, not a heap-use measurement.

The hypothesis is that early empty-state selection contributes to the opaque
construction/drop cost and early syntax regressions observed in 0802. This
single source intervention includes its compiler and code-layout effects. It
does not isolate a particular instruction or layout feature as the cause.

## Protocol and correctness

The [packet](results/change-0803/README.md) retains 39 literal cases, two modes,
six alternating paired native blocks, 30 samples, three warmups and 4,096
iterations per child, pinned to CPU 12. Ratios use nearest-rank process p50s
and a median bootstrap with 10,000 resamples and fresh seed 803080. No
historical native timing is pooled. A diagnostic regression means a median
ratio above 1.05 with the 95% interval wholly above 1; every flag is retained.
The decision always leaves advancement and adoption false.

Construction measures the common dispatch and an opaque iterator reference
through `black_box(&iterator)`, drop and checksum work. Consumption also
includes construction and checksum work. These hot repeated micro-inputs
provide no public-workflow latency estimate. Two separate Callgrind repeats
retain guest instruction and branch diagnostics, not native timing, hardware
branch rates, allocator API counts or phase fractions.

All five helper copies pass the exact isolated mirror tests: 100 before and
100 after, with warnings-denied Clippy for both. The release probe passes
formatting, build, check, Clippy, independent fixture verification and semantic
self-check across all 39 cases and clone advances 0, 1, 2, 3, 4, 5, 32 and 33.
These are helper/probe checks, not full production-workspace verification.

The initial after test failed because it asserted the former private `Done`
state immediately after empty construction. Its retained source and receipts
show 19 passing tests and that one failure in the first after mirror. Only
that test is amended: `First` before the first request, `Done` after `None`,
and clone/repeated exhaustion parity. Iterator results stay equivalent, while
derived debug output before the first request can differ. All other tests and
the measured implementation remain unchanged by this amendment. The failed
attempt is retained; no failed build executable or capture exists.

## Results

All 39 construction rows improve against the rejected control, with median
after/before ratios from 0.296307 to 0.372513 and all intervals wholly below 1.
No construction or consumption row triggers the frozen diagnostic regression
flag. The full 78-row table remains in [summary.md](results/change-0803/summary.md).

| Case and mode | After/before p50 | Change | 95% interval |
| --- | ---: | ---: | ---: |
| distinct-0/construct | 0.314006 | -68.599% | 0.296007–0.338071 |
| distinct-0/consume | 0.519075 | -48.093% | 0.518323–0.527919 |
| distinct-1/consume | 0.761308 | -23.869% | 0.761008–0.768881 |
| distinct-2/consume | 0.868048 | -13.195% | 0.863809–0.873600 |
| distinct-4/consume | 0.933885 | -6.612% | 0.910526–0.954013 |
| distinct-32/consume | 0.999337 | -0.066% | 0.984156–1.009495 |
| syntax-flag-after-0/consume | 0.814159 | -18.584% | 0.780290–0.816137 |
| syntax-equals-value-after-2/consume | 0.907797 | -9.220% | 0.894010–0.911018 |

The largest consumption median increase is 0.558% for an unquoted duplicate
after 33 attributes. Three consumption medians exceed 1; none meets the
diagnostic regression threshold. There are 73 spread flags among 156
case/mode/leg groups, including construction variability. Those flags and all
six paired ratios remain visible; the intervals above do not hide them.

Both Callgrind repeats agree on the following guest instruction counts.
All 312 owners qualify; all 624 positive/termination dumps conserve all five
event counters, and termination dumps are zero.

| Case and mode | Before Ir | After Ir |
| --- | ---: | ---: |
| distinct-0/construct | 95 | 70 |
| distinct-0/consume | 134 | 122 |
| distinct-1/consume | 424 | 407 |
| distinct-2/consume | 887 | 870 |
| distinct-4/consume | 2,409 | 2,390 |
| distinct-32/consume | 35,997 | 35,978 |
| syntax-flag-after-0/consume | 291 | 274 |
| syntax-equals-value-after-2/consume | 1,289 | 1,270 |

The relocation materially changes construction cost and early consumption in
this binary. Middle-size consumption retains almost the same guest instruction
work: 32 attributes change by only 19 instructions. This supports keeping
constructor placement separate from the still-unresolved bounded linear
comparison cost. It does not identify a unique native-time cause, and multiplying
these ratios by historical 0802 ratios would not create a production comparison.

## Disposition and remaining work

The packet retains 936 native reports with 28,080 samples and 312 profile
reports with one sample each: 1,248 reports and 28,392 samples. Independent
raw-statistic, checksum, semantic and profile-conservation audits reproduce the
results. The decision retains `advance_to_workflow_trials: false` and
`production_adoption: false`.

Before proposing another combined candidate, independently measure the linear
stage’s ordering-comparison equality check while holding this placement fixed.
A resulting candidate would still need a fresh comparison against current
production, followed by workflow, resource and cross-format qualification.

All 9,196 production source hashes, 35 architecture inputs, unrelated files
and other worktrees remain unchanged. No CRUD, memory, cold/range, producer
or concurrency coverage is promoted. iWork is excluded; the broader GOAL
remains incomplete.

The owned target was removed after executable identity verification, freeing
178,371,844 logical bytes. Cleanup retains the exact final executable identity;
no failed-build executable existed. Post-cleanup replay and final sealing
verify the retained evidence and unchanged production sources.
