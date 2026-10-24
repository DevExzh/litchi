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

Before the runtime slot opened, the source-only checks below passed:

```text
rustfmt --edition 2024 <all changed Rust files>
git diff --check
```

No release artifact, benchmark binary, native timing, or large evidence
artifact is claimed by this record.

## Follow-up consumer correction

The standalone `tools/perf-baseline/unified-batch-closure` Keynote leaf leg was
also migrated: its `edit.text()` read now propagates the deferred
`litchi_keynote` error with `?` before computing the UTF-16 replacement span.
The remaining `docs/unified-ops/results/0359a/routes_before.rs` occurrence is
historical evidence for the pre-API source-backed baseline and is not compiled
or used by the current harness.

## Runtime validation receipt

The successful runtime checks below were run from source `88869b24b9826cb8cdf45e8d7cd0e3ebebae9852` with Rust `1.95.0`,
the assigned offline/locked `CARGO_BUILD_JOBS=1` environment, and the
workspace's sparse patterns unchanged. Raw logs remain outside the repository
under `/home/zhuhe/code/litchi-worktrees/.coord/keynote-lazy-entry/`:

```text
cargo test --offline --locked -p litchi-keynote --all-features -j 1
  receipt-keynote-test.log (121384 bytes, sha256 e7805ac601b6643bca72cd5831300a17e0b51283f22e5498dbe4ff2ee3d3fc6f)
  77 test targets; 1169 tests scheduled; 1168 passed and 1 ignored; 0 failed.

cargo clippy --offline --locked -p litchi-keynote --all-targets --all-features -j 1 -- -D warnings
  receipt-keynote-clippy-v2.log (177 bytes, sha256 3d4a0ec6a3fd075b8af30abdbd3916b0b85e86488488249cc146ba53314e8121)
  exit 0.

cargo clippy --offline --locked -p litchi-iwa --all-targets --all-features -j 1 -- -D warnings
  receipt-iwa-clippy.log (866 bytes, sha256 4ecf9794bd173ef8534d826c8be2d7905776aa9e001056efa315cfcc4f6df3d6)
  exit 0.

cargo check --offline --locked -p litchi-ops --features keynote --all-targets -j 1
  receipt-ops-keynote-check.log (2872 bytes, sha256 5dd426ea858a82a7d09b151f53a66ab4452b9b55119ce6b47ca490762cc0aadf)
  exit 0.

cargo clippy --offline --locked -p litchi-ops --features keynote --all-targets -j 1 -- -D warnings
  receipt-ops-keynote-clippy.log (1487 bytes, sha256 e809d145eb256dca48f8b15aaebf9a208507c9c7f1fcaf693badad6d22fe82a8)
  exit 0.

cargo check --offline --locked --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --no-default-features --features keynote --all-targets -j 1
  receipt-batch-harness-check.log (3357 bytes, sha256 c78e08606ddf9505dd6a29f774089353cf5250ecb9347b6076c614a425bf2b6c)
  exit 0.

cargo clippy --offline --locked --manifest-path tools/perf-baseline/unified-batch-closure/Cargo.toml --no-default-features --features keynote --all-targets -j 1 -- -D warnings
  receipt-batch-harness-clippy.log (212 bytes, sha256 f60b6f6b2ebf065a11e072a6caa587ffd8743213edb149cc4420a065967fe712)
  exit 0.
```

The standalone harness lockfile was regenerated offline because its tracked
lock predated the current workspace dependency graph; the resulting
`tools/perf-baseline/unified-batch-closure/Cargo.lock` is included with this
consumer migration. The lock-generation log is external:
`receipt-batch-harness-lock.log` (333 bytes, sha256
`d9e77ab981e09f171390cc276310b53154d9a8e1c78508bcb1c3372a81ac49d4`).

The full `litchi-iwa` test run was also attempted with
`cargo test --offline --locked -p litchi-iwa --all-features -j 1` and reached
1695 passed / 1 failed. Its sole failure is the pre-existing assertion in
`package::tests::package_edit_rejects_limit_exceeding_stage_atomically`: the
test expects `total size` although the replacement payload first violates the
4-byte per-member limit and the implementation correctly reports `member`.
It is unrelated to Keynote and is not changed in this lane. Raw evidence is
external at `receipt-iwa-test.log` (171661 bytes, sha256
`cc2a859f1dd4b9f67327c76f791d316767581039806ca49e788cab95d356c0d2`).
