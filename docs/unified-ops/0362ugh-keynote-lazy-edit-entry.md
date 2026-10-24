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

## Baseline diagnosis and gate repair

The initial IWA result was not accepted as a baseline exception. A clean
detached checkout at `c3f56b7a9d39973a079f3bcdeb5fe2cd74a2deaa` reproduced the
same focused failure with Rust `1.95.0`, the same root `.cargo/config.toml`,
the same offline lock (`Cargo.lock`, 139592 bytes, sha256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`), and
the same all-feature command:

```text
cargo test --offline --locked -p litchi-iwa --all-features package::tests::package_edit_rejects_limit_exceeding_stage_atomically -- --nocapture
  receipt-iwa-baseline-focused.log (1033 bytes, sha256 182ed0a2761f2afab8adee25722d841c1e2f4871b9af3bc285ef18917d67beb7)
  test result: FAILED. 0 passed; 1 failed; 1695 filtered out.
  panic: assertion failed: error.to_string().contains("total size")
```

The failure was a test-fixture defect, not a production behavior to waive:
the test used both `max_entry_bytes = 4` and `max_total_bytes = 4`, then
expected a five-byte replacement to reach the aggregate check. The narrow
repair changes only that test's per-member ceiling to 5, retaining the
aggregate ceiling at 4. The five-byte replacement now exercises the intended
total-size refusal without weakening the assertion or changing production
code.

After the repair, focused and full IWA validation passed from source
`8f747b7f0abe39ae9b5f098b0832b9b7967c2340`:

```text
cargo test --offline --locked -p litchi-iwa --all-features package::tests::package_edit_rejects_limit_exceeding_stage_atomically -- --nocapture
  receipt-iwa-focused-repaired.log (5715 bytes, sha256 668e5c2b1b6d3ec191e83f2423ef2550ca9751b66479d59e5d20e065a9bd2a63)
  focused test: 1 passed; 0 failed; 1695 filtered out.

cargo test --offline --locked -p litchi-iwa --all-features -j 1
  receipt-iwa-test-repaired.log (183373 bytes, sha256 ce9dbbb8c21f43680c526d293751bb00a33eda777eaa575801473797f2c9eb82)
  1696 unit tests, all integration targets, and 8 doctests passed; 0 failed.

cargo clippy --offline --locked -p litchi-iwa --all-targets --all-features -j 1 -- -D warnings
  receipt-iwa-clippy-repaired.log (169 bytes, sha256 70e961835c0ace8172902a789c97d73c6148ff972f002fa858b6bd06234b049d)
  exit 0.
```

All raw logs remain outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/keynote-lazy-entry/`.
