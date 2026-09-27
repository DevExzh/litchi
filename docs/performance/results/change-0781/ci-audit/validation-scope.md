# CI matrix repair validation

The workflow's former constants required 37 default cases and 201 full rows.
The current default source and authoritative identity manifest require 41 cases
and 213 rows. The workflow also selected historical CRUD coverage index v1;
the active coverage contract is v2. The archived before-workflow and defect
record retain the original state.

The replacement checker derives expected rows from the manifest, checks its
identity digest, requires an exact matrix, and validates sample counts, units,
sample-order permutations, nonnegative integer elapsed samples, sink invariants,
and applicable identity configuration. Workflow paths, validation, copied index,
and uploaded artifacts now consistently select v2. Full measurement remains
schedule/manual-only; push and pull-request smoke scope is unchanged.

Root built the unchanged baseline harness at
`fdca3e63037ac83ff27d4bf64a0f157b971b45f8` offline with the retained lockfile.
The actual tiny/compressible smoke report has 41 rows and two samples per row;
the actual default full report has 213 rows and 15 samples per row. These are
local runtime contract checks, not timing comparisons or a claim that hosted
CI has run. This harness build precedes the proposed PPT production change.

The final local CI quality attempt currently retained in `ci-quality-1` passes
all six gates: 75 Python unit tests, smoke matrix validation, full matrix
validation, the active v2 coverage contract, full-report corpus binding, and
the v2 coverage contract bound to the full report and source selectors.
`ci-quality-0` retains the earlier successful 74-test attempt. A review observed
the helper during a concurrent numeric-sample fix; its reported boolean/NaN/
negative-value blocker is resolved in the frozen helper and exercised by the
tests. Unit, ordering, and configuration checks were strengthened afterward.

The helper does not independently recompute elapsed summary statistics. This
repair establishes the matrix and capture contract rather than statistical
comparison correctness. The manifest's historical `main.rs:Case::DEFAULT`
locator also remains a coordinated provenance follow-up; the actual default
declaration resides in `tools/perf-baseline/src/lib.rs`. No CRUD scenario is
promoted from correctness-only coverage by this repair, and no end-to-end
performance improvement is claimed for the CI changes.
