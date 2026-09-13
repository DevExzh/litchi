# OLE2 validated directory-name handoff

Status: candidate rejected by mandatory native performance gates; exact baseline source restored. All eight final quality commands passed. The matched evidence and rejection decision are retained in change-0554.

This experiment starts at `53330ff67` and tests private CFB directory-name ownership transfer. The validation pass already decodes each normal directory name; the candidate moves that validated string into the public entry instead of decoding and allocating it again. Root decoding stays separate to preserve the supported classic-Mac root's historical public name and canonical lookup name. Scalar parsing, graph validation, resource limits, and final publication remain in their existing phases.

Fresh baseline constructor profiles show 257 public name-decoder calls per many-small CFB open, 5 for few-large, 4 for tiny, and 12 for the XLS owned-source one-cell case. These are incoming decoder calls in positive timed-owner dumps, not allocator calls or an inferred directory count from a manifest. The candidate prediction is one root decode per open. Native latency and memory results remain separate from instruction attribution.

The [frozen plan](results/change-0554/plan.json) compares nine XLS workflows and three CFB shapes in two native ABBA repeats with 1,000 samples and 20 warmups. Matched allocator captures use 30 samples and three warmups. Constructor profiles retain five timed samples per workload and repeat, plus setup and termination dumps. The [admission supplement](results/change-0554/admission-supplement.json), recorded before any captures, preserves the prior requirement for at least 3% median improvement in each of four primary XLS workflows in both repeats. Many-small CFB p50 and mean must also improve by at least 3%; native controls and process peak RSS must remain within the 5% ceiling. No single metric overrides another failed mandatory gate.

The selected candidate passed 314 CFB tests, warning-denied Clippy and formatting. One existing CLSID documentation example remains ignored. Focused tests compare public scalar fields and name-cache identity, the classic-Mac two-view fixture, exact malformed-name errors, and failed-load publication. The [source review](results/change-0554/candidate-review.md) and [resource-boundary review](results/change-0554/resource-boundary-review.md) distinguish removal of a duplicate allocation from any claim of allocator-failure schedule equivalence.

OLE2 and OOXML performance remain the priority. ODF is deferred until that goal completes; iWork is excluded.

## Matched outcome

Both stages completed all native, allocation and profile captures. The numerical analyzer validates the full matched matrix. Many-small CFB median latency improved 9.41% and 12.99% in the two repeats; its mean improved 9.60% and 12.74%. Allocation calls fell 47.06%, allocated bytes 12.43%, and incremental region peak 13.79% for that shape.

The four primary XLS workflows all regressed in both repeats:

| XLS workflow | Median change R1 | Median change R2 |
| --- | ---: | ---: |
| Source-backed open | +10.93% | +2.31% |
| Source-backed open + one cell | +7.10% | +5.02% |
| Owned-source open | +8.91% | +5.21% |
| Owned-source open + one cell | +8.18% | +6.54% |

The eight primary XLS improvement checks all failed. Nineteen of 52 native control checks exceeded the 5% ceiling; all 72 allocation checks and the process peak-RSS controls passed. The comparison therefore rejects the production candidate independently of the profile outcome. No threshold was relaxed and no timing was selectively rerun.

All 83 matched adverse rows and 43 same-build drift rows are retained for individual review. These diagnostics do not establish a cause for the XLS regression. Instruction counts and allocator regions cannot replace native end-to-end latency evidence. Measurements are retained in [metrics-analysis.json](results/change-0554/metrics-analysis.json).

Fresh XLS owner attribution identifies physical-sector reconciliation as the largest independent leaf beyond the already-investigated collector. The [next attribution note](results/change-0554/next-attribution.md) proposes proof-first physical marker accounting; it supplies no candidate performance result and excludes the earlier rejected paired-prefix approach.

Final-source checks passed on the exact restored baseline: 4,228 tests in 154 groups (27 existing ignored examples/tests), workspace all-features check, warning-denied legacy-format Clippy and rustdoc, formatting, crate boundaries, no-default-feature checks and strict performance-claim checks. Candidate code remains only in the experiment snapshot.

The [matched profile comparison](results/change-0554/profile-comparison.json) and [instruction mapping](results/change-0554/instruction-comparison.json) pass. Every candidate timed open retains exactly one root decoder call; many-small decoder self instruction Ir fell 99.835% across the five timed dumps per repeat, while CLSID calls and mapped self instruction work are preserved. Exact raw positions map to the retained symbol-bounded disassembly. These findings establish the attempted mechanism; they do not explain or override the native XLS regressions.
