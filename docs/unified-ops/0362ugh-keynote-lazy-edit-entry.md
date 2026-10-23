# 0362ugh — Keynote lazy edit entry

This lane keeps Keynote edit construction O(1) with respect to package size.
It starts from the W7 API head `c3f56b7a9d` on `uo2/keynote-lazy-entry` and
does not change the read-only projection or patch/replay contracts.

## Deferred edit selection

The Keynote text, chart arrangement, chart title, chart data, table appearance,
and table lock-state edits now retain only their immutable package reference,
owned archive-free selectors, and compact staged state at entry. Slide names
and chart names are copied as bounded caller input; package graph scans,
selected payload decoding, and text/chart/table snapshots are deferred until a
value is read, a staging operation requires the source value, or `commit()`
runs. `OnceLock` caches each successful selection/snapshot for subsequent
accesses without moving the work into package open or hiding it in drop.

Chart-arrangement aspect setters retain their fluent, infallible API: a lock or
aspect request is recorded as a compact flag and applied to the source
arrangement at read/commit time. Complete arrangement replacement supersedes
those flags, matching the prior setter ordering semantics.

The deferred paths preserve the existing native ownership checks, typed
refusals, limits, exact no-op commits, source-bound patches, candidate
verification, and redacted debug output. Because selector and value accessors
can now be the first operation that discovers malformed or missing source
content, their consumers were migrated to handle the family error type. Invalid
selector/source expectations in Keynote, iWork editor, root-facade, and fuzz
consumers now force the edit through `commit()` before asserting failure.

Focused tests assert an unmaterialized debug state immediately after entry,
materialization after the first read, and a malformed chart edit that can be
constructed and dropped without decoding before commit is attempted. Existing
exactness, refusal, replay, and limit tests remain on the same commit paths.

## Static validation

The assigned lane was source-only while the coordinator held Cargo slots. The
following checks passed:

```text
rustfmt --edition 2024 <all changed Rust files>
git diff --check
```

No Cargo build, test, release artifact, benchmark binary, native timing, or
large evidence artifact is claimed by this record.

## Follow-up consumer correction

The standalone `tools/perf-baseline/unified-batch-closure` Keynote leaf leg was
also migrated: its `edit.text()` read now propagates the deferred
`litchi_keynote` error with `?` before computing the UTF-16 replacement span.
The remaining `docs/unified-ops/results/0359a/routes_before.rs` occurrence is
historical evidence for the pre-API source-backed baseline and is not compiled
or used by the current harness.
