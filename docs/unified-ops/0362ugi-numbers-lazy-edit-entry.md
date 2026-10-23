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

The chart-data edit keeps optional bounded selection and `Arc<ChartData>` state
and loads it only for `before`, `after`, or `commit`.  Image adjustments retain the
sheet/image positions and load the graph selection only when a value is
requested or published.  Table-title settings retain the selected positions;
its target and decoded settings are loaded by `settings` or `commit`.
Table-merges keeps the selected positions and budget until `regions`, a merge
operation, or `commit`; its source regions and dimensions are then decoded
once and reused by the existing mutation and exact-publication paths.

The iWork Numbers chart consumer now calls the fallible chart `before` getter
while it still owns the requested grid.  A deferred `UnsupportedDependency`
therefore transfers that same grid to the legacy full replacement route; a
successful focused edit moves the grid once into `set` without an unconditional
full-grid clone.

No panic or placeholder value is used for a deferred failure.  Existing typed
source, refusal, budget, verification, patch-conflict, exact-no-op, and
replay behavior remains on the fallible operation that now performs the
deferred read.  Public getters whose value can now fail return the family's
existing error type, and the in-workspace table-merges test is migrated to the
fallible `regions` accessor.

Validation completed in source commit
`468bfb29d13f5827dc143d505ec5ce53f2314468` with Rust 1.95.0, the assigned
one-job capped environment from `.coord/numbers-lazy-entry/env.sh`,
`--offline --locked`, and an external target directory (4,733,132,163 bytes
at receipt time).  `Cargo.toml` SHA-256 is
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`; the
matching external `Cargo.lock` SHA-256 is
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.

The raw logs remain outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/`:

* `full-numbers-tests-final.log` — 118,554 bytes,
  SHA-256 `8c402e6bbfec17fe8f7e4da33a5e704d97dbfe923957765ef6ea360ee56e7e0f`.
  Command: `source /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/env.sh && set -o pipefail && cargo test --offline --locked -p litchi-numbers --all-features --all-targets --no-fail-fast --message-format=short 2>&1 | tee /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/full-numbers-tests-final.log`.  It reported 86 suites, 1,077 passed, 0 failed, and 4 ignored.
* `numbers-clippy-v2.log` — 177 bytes,
  SHA-256 `45eed02fe309724e174579852fcb08b982d57187f477b4b89141f27cbd5d4c9a`.
  Command: `source /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/env.sh && set -o pipefail && cargo clippy --offline --locked -p litchi-numbers --all-targets --all-features -- -D warnings 2>&1 | tee /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/numbers-clippy-v2.log`; it exited 0 with no diagnostics.
* `iwa-chart-fallback-test-v2.log` — 536 bytes,
  SHA-256 `cf50714a1ddb14a1b8a54497f28a8839c6eb51df4cd352d4f8e228fde0b8024d`.
  Command: `source /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/env.sh && set -o pipefail && cargo test --offline --locked -p litchi-iwa --all-features --lib deferred_chart_dependency_refusal_keeps_full_replace_input --no-fail-fast --message-format=short 2>&1 | tee /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/iwa-chart-fallback-test-v2.log`; it reported 1 passed, 0 failed, 1,696 filtered.
* `numbers-facade-test.log` — 17,842 bytes,
  SHA-256 `0e5a23a1c2bbeaf2517caf5b397dd663313df818457878ad7a6f73ac4cdfd059`.
  Command: `source /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/env.sh && set -o pipefail && cargo test --offline --locked -p litchi --all-features --test numbers_facade --no-fail-fast --message-format=short 2>&1 | tee /home/zhuhe/code/litchi-worktrees/.coord/numbers-lazy-entry/numbers-facade-test.log`; it reported all 5 facade tests passed and 0 failed.

`cargo fmt --all -- --check` and `git diff --check` also passed.  The source
commit is ready for coordinator integration.
