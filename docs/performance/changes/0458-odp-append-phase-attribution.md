# 0458 — ordinary ODP append phase attribution

The ordinary append baseline previously exposed only total lifecycle cost.
A standalone harness command now measures snapshot opening, transaction
construction, append, commit and sequential output at existing public API
boundaries, alongside an unsegmented comparison. Production behavior and
registry counts remain unchanged.

The two-repeat, three-shape normal/allocator matrix passes 24 lanes and 720
samples. Large normal commit shares are 45.64 / 45.55%; snapshot opening and
transaction creation together take about 53%. Commit allocates 118,096,476 of
211,442,207 lifecycle bytes. Phase allocation sums exactly equal direct totals.
No timing-quantile or maximum-RSS instrumentation comparison exceeds an
absolute 5% difference. These are diagnostics, not production speedup claims.

The full release harness suite passes 387 tests with one ignored. Strict
Clippy, build, formatting, warning-denied rustdoc, crate boundaries and eighteen
oracle checks pass. Two separate profiles pass, but all 1,656 sampled stacks
lack phase-marker ancestry, leaving internal-stage attribution open.

The [bundle](../results/change-0458/README.md) retains frozen inputs, raw rows,
profiles, recomputed summaries, portable verification and owned staging cleanup.
The full non-iWork goal remains open.
