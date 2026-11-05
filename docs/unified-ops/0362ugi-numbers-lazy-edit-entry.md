# 0362ugi — Numbers lazy edit entry

This slice addresses the confirmed edit-construction cost debt in the Numbers
title, chart-data, table-merges, and image-adjustment families.  It is based on
coordinator integration `c3f56b7a9d`; no native qualification is claimed.

The edit constructors retain only the immutable package reference, owned
selectors, cheap syntax/options checks, and the operation budget.  Name
catalog resolution, native target resolution, selected payload reads, wire
decoding, semantic value copies, and merge-region clones move to the first
value request, mutation, or commit.  Dropping an untouched edit therefore
does not scan a workbook catalog or copy a chart grid, image graph, title
payload, or merge-region vector.

The chart-data edit keeps an owned sheet selector, optional bounded selection,
and `Arc<ChartData>` state and loads them only for a position/data getter,
`before`, `after`, or `commit`.  Image adjustments retain an owned sheet
selector and image position and load the graph selection only when a position,
value, mutation, or publication is requested.  Table-title settings retain
owned sheet/table selectors and resolve positions, target, and decoded settings
only for `path`, `settings`, or `commit`.  Table-merges keeps owned selectors,
optional positions, and its budget until `regions`, a merge operation, or
`commit`; its source regions and dimensions are then decoded once and reused
by the existing mutation and exact-publication paths.

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
one-job capped environment from `(deleted-coord-artifact)/env.sh`,
`--offline --locked`, and an external target directory (4,733,132,163 bytes
at receipt time).  `Cargo.toml` SHA-256 is
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`; the
matching external `Cargo.lock` SHA-256 is
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.

The raw logs remain outside the repository at
`(deleted-coord-artifact)`:

* `full-numbers-tests-final.log` — 118,554 bytes,
  SHA-256 `8c402e6bbfec17fe8f7e4da33a5e704d97dbfe923957765ef6ea360ee56e7e0f`.
  Command: `source (deleted-coord-artifact)/env.sh && set -o pipefail && cargo test --offline --locked -p litchi-numbers --all-features --all-targets --no-fail-fast --message-format=short 2>&1 | tee (deleted-coord-artifact)/full-numbers-tests-final.log`.  It reported 86 suites, 1,077 passed, 0 failed, and 4 ignored.
* `numbers-clippy-v2.log` — 177 bytes,
  SHA-256 `45eed02fe309724e174579852fcb08b982d57187f477b4b89141f27cbd5d4c9a`.
  Command: `source (deleted-coord-artifact)/env.sh && set -o pipefail && cargo clippy --offline --locked -p litchi-numbers --all-targets --all-features -- -D warnings 2>&1 | tee (deleted-coord-artifact)/numbers-clippy-v2.log`; it exited 0 with no diagnostics.
* `iwa-chart-fallback-test-v2.log` — 536 bytes,
  SHA-256 `cf50714a1ddb14a1b8a54497f28a8839c6eb51df4cd352d4f8e228fde0b8024d`.
  Command: `source (deleted-coord-artifact)/env.sh && set -o pipefail && cargo test --offline --locked -p litchi-iwa --all-features --lib deferred_chart_dependency_refusal_keeps_full_replace_input --no-fail-fast --message-format=short 2>&1 | tee (deleted-coord-artifact)/iwa-chart-fallback-test-v2.log`; it reported 1 passed, 0 failed, 1,696 filtered.
* `numbers-facade-test.log` — 17,842 bytes,
  SHA-256 `0e5a23a1c2bbeaf2517caf5b397dd663313df818457878ad7a6f73ac4cdfd059`.
  Command: `source (deleted-coord-artifact)/env.sh && set -o pipefail && cargo test --offline --locked -p litchi --all-features --test numbers_facade --no-fail-fast --message-format=short 2>&1 | tee (deleted-coord-artifact)/numbers-facade-test.log`; it reported all 5 facade tests passed and 0 failed.

`cargo fmt --all -- --check` and `git diff --check` passed for the validated
commit above.  The selector-deferral successor `86c762d2a7284e3d45e2168f86865f13ea9fe3af`
was reviewed and extended in source commit
`4c2fd76b49` to admit owned selector names against each family's retained or
transaction ledger before copying, charge chart/image first-use sheet scans,
and migrate the four standalone fuzz `path()` consumers to the fallible
accessor.  A follow-up source commit `6a053c36d3` restores each retained
edit budget after deferred selection, payload, or clone errors instead of
turning the next accessor into a generic `InvalidSource`; it also exercises
repeated getter/mutation/commit calls and proves cumulative transaction-work
accounting after a typed selector refusal.  Rustfmt and `git diff --check`
pass for these successors.  Cargo validation was not rerun in this turn
because the coordinator's build slot was transferred to another workstream;
rerun the Numbers, iWork, facade, and strict-Clippy commands above on
`6a053c36d3` before coordinator integration.
