# 0362ugi — Numbers lazy edit entry

This slice addresses the confirmed edit-construction cost debt in the Numbers
title, chart-data, table-merges, and image-adjustment families.  It is based on
coordinator integration `c3f56b7a9d`; no native qualification is claimed.

The edit constructors retain only the immutable package reference, checked
semantic positions, selectors already resolved through the package's cheap
projection, and the operation budget.  Native target resolution, selected
payload reads, wire decoding, semantic value copies, and merge-region clones
move to the first value request, mutation, or commit.  Dropping an untouched
edit therefore does not copy a chart grid, image graph, title payload, or
merge-region vector.

The chart-data edit keeps a bounded `OnceCell` for its `Arc<ChartData>` and
loads it only for `before`, `after`, or `commit`.  Image adjustments retain the
sheet/image positions and load the graph selection only when a value is
requested or published.  Table-title settings retain the selected positions;
its target and decoded settings are loaded by `settings` or `commit`.
Table-merges keeps the selected positions and budget until `regions`, a merge
operation, or `commit`; its source regions and dimensions are then decoded
once and reused by the existing mutation and exact-publication paths.

No panic or placeholder value is used for a deferred failure.  Existing typed
source, refusal, budget, verification, patch-conflict, exact-no-op, and
replay behavior remains on the fallible operation that now performs the
deferred read.  Public getters whose value can now fail return the family's
existing error type, and the in-workspace table-merges test is migrated to the
fallible `regions` accessor.

Validation is pending the coordinator's assigned one-job Cargo slot.  Before
acceptance, run the full `litchi-numbers` test suite, affected facade and iWork
consumer tests, and strict Clippy with the capped offline/locked environment;
record exact counts and log hashes in a follow-up receipt.  The source is
formatted and `git diff --check` is clean.
