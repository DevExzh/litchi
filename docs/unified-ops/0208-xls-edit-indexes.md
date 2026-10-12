# 0208 — XLS edit inventory and staging indexes

The native XLS edit inventory now retains immutable coordinate indexes beside
its source-order entries. Cell lookup, duplicate detection, mutation admission,
semantic replay and candidate readback use the same indexes. Public cell
iteration and physical source slots keep their original order. Fixed numeric
candidate reuse shares the coordinate indexes because those edits cannot move
cells; structural candidates rebuild them from the admitted output.

Fixed and structural staging keep sorted sidecar indexes. They accelerate
coordinate and overlap searches while retaining the canonical operation
vectors, deterministic patch order and the 4,096-operation bound. Staging and
join reserve index capacity before changing logical transaction state. Joins
rebuild sidecars after sorting. Joining an effective change into an existing
no-op now reuses that coordinate's slot, so subsequent staging cannot leave a
stale duplicate operation. Joins also reject fixed/structural overlaps in both
directions.

Canonical `litchi_xls::Edit::source_cell` exposes an original-source lookup from
the edit's already admitted inventory. It deliberately reports the pre-edit
cell after staging and retains source-freshness checks around lookup, including
failures. Unified coalesced commands resolve their paths and use this native
lookup for presence, formula classification and scalar values. This avoids
repeated source payload decoding without adding a second facade cell model.

The indexes use fallible vector reservation. Lookup is logarithmic, but staging
an out-of-order coordinate can shift a vector tail; this is not a general
logarithmic-insertion or linear-scaling claim. The paired benchmark now accepts
`LITCHI_XLS_BENCH_REVERSE=1` to measure descending coordinate order as well as
its default ascending order. The index memory cost is retained with the native
snapshot; peak-memory qualification remains open.

## Verification

The full native suite passes 1,791 tests and doctests, including the public
source-cell example and eager/source-backed lookup tests. Lookups perform no
additional source payload reads, report original values after staging and
preserve stale-source precedence even for missing worksheets. The all-feature
facade suite passes 1,548 tests and doctests, including coalesced XLS barriers,
limits, error attribution, replay and exact inverses.

The native regressions cover duplicate refusal, source-order iteration,
no-op joins followed by staging, fixed/structural conflicts in both join
directions, and two-sheet out-of-order edits followed by join, clone and
further staging. Fixed and structural results and exact inverse bytes are
checked. All six focused index tests pass after the final two-sheet regression
was added.

Warning-denied all-target Clippy and rustdoc pass for XLS and the facade.
Minimal and XLS-only feature checks and the XLS glossary audit pass (24 banned
names and one transaction family). The crate-boundary audit passes for 66
packages and 268 internal dependencies with 11 existing debt items. The
OfficeCLI generated table remains current and its 15 generator tests pass.
Both repository lockfiles are unchanged. Offline checks use an external
lockfile selecting cached cc 1.6.0 and an external build target.

Reproduce with `cargo test -p litchi-xls --all-features`,
`cargo test -p litchi-ops --all-features`, and the focused native
`cell_values::index_tests` module. No Python/CLI or complete acceptance claim
is made.

## Remaining scope

Source-backed freshness cost, out-of-order insertion cost, resource/formula
staging, repeated-coordinate coalescing, source query streaming and broad
performance qualification remain follow-ups. Worksheet and axis CRUD,
formula/style editing, security publication, other formats, Python/CLI and
the complete goal remain open.

The base remains an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no additional rebase was needed.
