# 0207 — XLS cell transaction coalescing

Consecutive direct scalar cell sets, clears and removals in `File::edit` and
atomic `File::batch` now share a native XLS package edit. Distinct targets may
span worksheets. The facade retains command plans and attribution, while the
native owner retains all cell state, validation, dependency closure and
publication. No parallel document model or implicit runtime is introduced.

Repeated coordinates close the run and resolve again against the published
candidate. This preserves insert/remove/set sequences. Reads, ranges, queries,
worksheet renames, empty property maps and formula handling remain barriers.
Per-item `changed` flags distinguish scalar no-ops from mutations. Best-effort
batches retain their existing independent-item execution and rollback policy.

Deferred native failures are located by replaying bounded command prefixes
only on the failure path. A failed staging operation first closes any earlier
prefix so a preceding deferred error keeps its original item index. Source,
signature and encryption admission remain in force throughout the run.

## Bounds and native run boundaries

The native XLS owner has a hard limit of 4,096 staged changes. The first
10,000-operation measurement exposed that limit. A typed native operation
limit now closes a nonempty run, publishes it privately, and retries the same
command against the new candidate. A limit on an empty run remains an error.
Other native errors and caller limits are not converted into run boundaries.
This preserves support for larger public batches without silently widening
the native policy. The 1,000-operation scenario uses one native commit;
10,000 distinct numeric operations use three bounded native commits.

The public mutation budget counts calls, including no-ops. XLS previously
compared that bound with native cell/resource counts at publication, so a
text setter could consume multiple public operations. Publication now checks
the public command count and retains the native count separately in diagnostic
history. Staged-payload and retained-snapshot bounds remain independent.

## Verification

The new suite demonstrably fails on the preceding implementation's commit
counts: five distinct mutations produced five commits, and 1,000 produced
1,000. With coalescing they each produce one. A new 4,100-operation boundary
test passes with two commits for both sets and removals, through eager and
real filesystem owners. The original source and exact inverse are checked.

The all-feature facade suite passes 1,547 tests and doctests. Following the
final mutation-budget correction, all 49 focused XLS write/removal/batch tests
pass, including 13 batch tests. Coverage includes cross-sheet runs, repeated
coordinates, query/read/rename barriers, no-op item flags, atomic/best-effort
attribution, a deferred rich SST dependency refusal, source freshness,
signed/encrypted admission, budgets, exact inverse and replay.

Warning-denied all-target Clippy and rustdoc pass. Minimal and XLS-only
feature checks pass, as does the XLS glossary audit (24 banned names and one
transaction family). Scoped formatting passes. The crate-boundary checker passes
for 66 packages and 268 internal dependencies (11 existing debt items).
The generated OfficeCLI table remains current, and all 15 generator tests
pass. No capability claims or property registry entries changed. Repository
lockfiles remain unchanged; offline validation uses an external lockfile
selecting cached cc 1.6.0 and an external build target.

Reproduce with `cargo test -p litchi-ops --all-features` and the focused
`xls_unified_batch`, `xls_unified_write`, and `xls_unified_remove` integration
tests. No Python/CLI or full acceptance completion is claimed.

## Measurement method

`crates/litchi-ops/examples/xls_batch_bench.rs` pairs native package edits
with both unified closure and atomic batch APIs for 1, 10, 100, 1,000 and
10,000 distinct numeric sets and removals. Eager and filesystem owners use
the same input bytes. Above the native staging limit, all legs use equivalent
bounded transactions and assert equal commit counts. Generation, opening,
payload construction, serialization and correctness verification are outside
timing. Native whole, staging (including edit admission), and commit costs
are reported separately. Every output checks complete coordinates/values,
byte equality and source immutability.

## Remaining scope

The native inventory lookup and staged-change overlap/search paths still
contain linear scans per command. Their indexing, source-backed direct
lookup overhead, repeated-coordinate coalescing and broad batch scaling
qualification remain follow-ups. Worksheet and axis CRUD, formula/style
editing, security publication, source query streaming, the remaining formats,
Python/CLI and the complete goal remain open.

The base remains an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no additional rebase was needed.
