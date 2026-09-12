# 0529: reject the XML attribute probe under native timing gates

The zero/one-attribute probe is rejected and the shared XML auditor is restored
exactly to baseline. Publication allocation calls fell substantially, but the
frozen native timing requirements did not pass. This batch retains separate
publication allocation instrumentation, four compatible public XML regression
tests, and all candidate evidence. It claims no retained runtime speedup.

The candidate followed the publication attribution in [0528](0528-xlsx-publication-attribution.md).
It probed at most two unchecked attribute results before any state mutation.
A proven zero/one-attribute tag avoided duplicate-key bookkeeping; every error
or second result replayed the original checked iterator. Both slice and
streaming audit paths retained their validation, limits and error ordering.
No validation pass, public API, dependency or unsafe boundary changed.

The measurement harness adds `publication_allocation_metrics` alongside the
existing commit vector. The region includes publication and returned snapshot
destruction, matching native timing. Normal binaries report explicit
unavailable values; allocator binaries provide independent measured samples.
The harness is identical in both source-bound builds.

The frozen pilot used medium and dense-sparse XLSX one-percent edits, two
repeats of 200 measured samples after 20 warmups, seven guards per repeat,
and five allocator samples per shape/repeat. Native order was baseline A1,
candidate B1/B2, retained baseline A2. All 36 native children and eight
allocator children passed their output/lifecycle checks: 2,440 native durations,
40 publication allocation samples and 40 separate commit diagnostics.

Positive percentages below mean reduction; each row required total p50 and
mean reductions of at least 2% and publication p50 reduction of at least 5%.

| Shape | Repeat | Total p50 reduction | Total mean reduction | Publication p50 reduction |
| --- | --- | ---: | ---: | ---: |
| dense-sparse | 1 | 0.5373% | 0.4985% | -0.0823% |
| medium | 1 | -0.7667% | -0.9106% | -0.0163% |
| dense-sparse | 2 | 1.5896% | 1.4953% | 5.1574% |
| medium | 2 | 2.1301% | 1.3717% | 8.5325% |

Every total-mean row fails. The first publication repeat is essentially flat
for both shapes; second-repeat improvements do not override those failures.
Publication allocation calls fall from 19,246 to 414 for medium and from
36,622 to 414 for dense-sparse, reductions of 97.8489% and 98.8695%, identical
in both repeats. Reallocation calls remain 40; incremental publication peak
is unchanged. Allocation reduction alone does not satisfy the timing contract.

All 46 matched adverse flags and 76 same-build drift flags are retained.
The managed noncompact repeat-1 guard is adverse in total p50 by 5.3910%,
mean by 5.5029%, p95 by 6.4320% and p99 by 7.5750%. Other matched flags
cover open, publication, reopen and planning metrics. No over-5% matched RSS
or allocation-peak flag is present. Drift is not silently removed or used to
turn a failed row into a pass. Conditional profile, hardware and eager lanes
remain unmeasured because the pilot failed.

The candidate passed all 14 quality checks with 2,025 successful test
executions. The initial test compile failure was a discarded must-use Report;
a corrected assertion passed before candidate freeze. The private checked-path
oracle was removed with the rejected production helper. Final restored-source
quality passes all 14 checks with 2,024 successful executions. Evidence replay,
cleanup and sealing are recorded in the bundle.

The next OLE2/OOXML investigation should attribute XLSX edit-planning work,
a sizeable phase not yet isolated in this sequence, before selecting a new
candidate. Do not revive the rejected probe or the earlier scanner/arena
patches without a distinct mechanism and fresh evidence. ODF optimization
remains deferred until the OLE2/OOXML goal completes; iWork is excluded.

Evidence: [0529 bundle](../results/change-0529/README.md),
[comparison](../results/change-0529/comparison.json),
[numeric review](../results/change-0529/numeric-review.md).
