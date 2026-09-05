# Validation and limits

Rust 1.98.1 was used for serialized builds and checks. `validation.json`
records exact commands, timestamps, exit codes and source hashes for the first
validation pass. The final focused rerun after the style fix and added Python
negative cases is bound by `final-validation.json`.

| Check | Outcome | Raw log |
| --- | --- | --- |
| Counter, boundary and concurrency tests | 20 passed, including final rerun | `final-counter-tests.log.gz` |
| Operation aggregation tests | 24 passed | `aggregation-tests.log.gz` |
| Filesystem sample serialization | 1 passed | `filesystem-serialization.log.gz` |
| Actual global allocator wrapper tests | 5 passed | `allocator-wrapper.log.gz` |
| Comparator compatibility and negative fixtures | Final 98 passed | `final-comparator-tests.log.gz` |
| Strict all-feature/all-target Clippy | Failed; existing harness debt plus one new collapsible-if diagnostic on first pass | `clippy-strict.log.gz` |
| Diagnostic Clippy after fixing the new warning | Exited successfully; no counter-file diagnostic, existing warnings remain | `final-clippy-diagnostic.log.gz` |
| Rustdoc | Passed with warnings denied, library/allocator feature/no dependencies | `rustdoc.log.gz` |

The first Python pass had 97 tests. Follow-up adds raw filesystem region-peak
negative fixtures and explicit V2/V3 identity mismatch coverage; the final
suite has 98 tests. Negative-fixture REGRESSION/INVALID output is expected.
The Rust follow-up only collapses an equivalent nested conditional in the new
peak updater; the final counter suite rechecks its behavior. The changed Rust files were processed with
`rustfmt --edition 2024 --config skip_children=true`. Unrelated existing
filesystem formatting was restored afterward; `checks/format-scope.json`
proves that normalizing the final five-line test-only change exactly reproduces
the tested source. No behavior or test logic changed in that restoration.

The diagnostic Clippy command allows the three established lints
`chunks_exact_to_as_chunks`, `clone_on_copy`, and `needless_lifetimes`.
It does not pretend to satisfy warning-denied policy: the final harness still
reports existing warnings. The broader goal's clean lint requirement remains
unmet. No lint allowance was added to source code.

The full document/harness test suite was not rerun for this observer-only
change. Focused tests exercise counter arithmetic, failure/poison/reentrancy,
region ownership and concurrent callbacks, vector alignment/status/cardinality,
filesystem serialization and the real System allocator wrapper. Fresh lifecycle
captures additionally exercise the observer through existing semantic,
preservation, refusal and output gates. Production document code is unchanged.
No new fuzz or native Office run is claimed by this batch.

The implementation revision is `40c40ea89591140e15e68d00f3fce2c4c376e445`. Source-hash receipts
retain the tested source boundaries, including the documented formatting-only
restoration. The clean candidate build binds the final source.

Repository checks passed: strict replay of all nine registered claims, crate
boundaries, and the non-iWork CRUD coverage index (15 categories, 30 selectors).
The index check is contract coverage, not a new timing report. Exact commands
and outputs are retained under `checks/repository-checks.json` and its logs.
Both release binaries built successfully from one clean detached checkout;
source identities before and after the build are identical. The normal report
omits the allocator revision and carries unavailable allocation vectors.

All four captures passed (120 measured observations), followed by nine actual-
report guard probes and three-check portable replay. The standalone bundle
test also rejects a changed pinned validator before replay. These checks pass
after removal of the candidate worktree and both copied binaries.

The initial derived summary was archived under `checks/initial-derived` before
clarifying table byte units, correcting an inherited V2 scope limitation, and
exposing the region invariant at summary level. Every per-run numeric and
identity record is identical; `checks/derived-revision.json` binds both versions.
An attempted overwrite was correctly refused and its log is retained. Raw
reports, journals and their verifier results were never changed.
