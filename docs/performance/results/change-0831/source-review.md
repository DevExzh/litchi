# Empty column-action owner map

The candidate adds only an empty-map return to `validate_column_actions`, after
the existing protected-sheet check. The exact before/after source is archived
under `candidate/`. No parser, range-assignment representation, writer or public
API changes.

The 0830 profile attributes one 524,288-byte allocation per real-file edit to
this validator's `Assignments<usize>` owner map. When `actions` is empty, no
iteration can query the map. Eliminating its construction also eliminates its
possible `"column interval map"` allocation failure for that empty-action path;
this is deliberate removal of unnecessary work, not a claim that an attempted
allocation can no longer fail.

The lossless source scan still precedes this validator. Malformed column XML,
invalid ranges, duplicate containers and source errors retain that validation.
Cell, row and defaults checks keep their order. Nonempty column actions retain
the existing protected-sheet and `StyleNeedsWidth` refusals, bounded owner map
and writer path. The two parser `Assignments<Properties>` allocations remain.

Two read-only reviews confirm this scope. Existing tests cover overlapping
column ownership, width/style retargeting and sparse splitting, protected-sheet
refusals, reversible patches and hidden shared-style identity. The pinned 0830
public probe additionally edits the real `dateAutofilter.xlsx` cell with a
stored column record and verifies the complete 0821 output bytes. It runs
fresh on both source legs. No redundant test that merely mirrors the guard is
added.

The nine-case performance matrix separates real-file public edit/lifecycle,
synthetic commit/save for one-cell and one-percent edits at three scales, and
a no-op control. Synthetic setup and edit planning occur before its commit/save
timer, unlike the real-file public edit interval. The no-op exits before this
validator and is a control, not an expected beneficiary. No benchmark in this
matrix performs a nonempty column action; existing unit/integration tests
provide that correctness coverage.

The root and standalone harness Cargo locks differ. Each remains pinned and
unchanged between legs, and both lock identities and their comparison are
retained. Workspace XLSX tests and the pinned oracle run under the root lock;
fresh harness library tests run under the benchmark lock. Preparation's first
authoring assertion incorrectly required these two existing locks to match;
that failure occurred before any freeze, build or capture and is retained.
