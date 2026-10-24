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

The retained selector names are admitted before ownership: each edit uses its
opened package's finite `semantic_limits().max_text_bytes()` as the selector
byte ceiling, reserves the exact UTF-8 copy fallibly, and reports the typed
`SelectorBytes` limit before constructing an `Arc<str>`. Chart edits charge
their slide and chart names cumulatively; positional selectors consume no
selector bytes. This check is local and O(1), so it does not restore the
package-wide show or chart-name scan to `edit_*()` entry.

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

The successful runtime checks below were run with Rust `1.95.0`, the assigned
offline/locked `CARGO_BUILD_JOBS=1` environment, and the workspace's sparse
patterns unchanged. They were run on three precisely recorded working-tree
overlays rather than one clean commit:

* `K` started at clean commit
  `88869b24b9826cb8cdf45e8d7cd0e3ebebae9852` and had five dirty, fallible
  getter-attribute files. Those files were committed unchanged by
  `8f747b7f0abe39ae9b5f098b0832b9b7967c2340` and remain unchanged in the
  final tree:
  `slide_chart_arrangement.rs` (`48987b4410b585c3468dba0d695befa3e8265763`),
  `slide_chart_data.rs` (`60a94b8ade7e4209c83cb8202d387e7d15be8cbb`),
  `slide_chart_title.rs` (`96d0d62cc1c0f9bac003a0ab7d44ddd784df0be4`),
  `slide_table_lock_state.rs` (`1404b54820c17c471611ac6f1c01765b27b16b56`),
  and `slide_text.rs` (`5a45920ca3c5254c1bbd5392cb293115bfb8623d`).
  The Keynote test/Clippy and ops checks below ran on `K`, not on clean
  `88869b24b9`.
* `H` was `K` plus the first offline-regenerated standalone harness lockfile,
  later committed by `8f747b7f0a` with blob
  `cc9461661455285e47a50d8ab08a573111a6f057`. The standalone harness checks
  below ran on that earlier `H`; they are retained as historical evidence, not
  as validation of the corrected lockfile now in this tree.
* `I` started at clean `8f747b7f0abe39ae9b5f098b0832b9b7967c2340` and had the
  one-line IWA test-fixture edit in `crates/litchi-iwa/src/package.rs`. That
  exact file is blob `f2834d280e9f5a3020848614c6e5f7a8c4fc5e6c` in final commit
  `178ef2ff3c68b9a26ff6146a0f71eed3ef405194`; the repaired IWA checks below
  ran on `I`, not on clean `8f747b7f0a`.

Commit `178ef2ff3c` differs from `8f747b7f0a` only in that IWA test
fixture and the receipt update then present. The Keynote, ops, and harness
source paths are byte-identical between those commits. The harness lock was
then corrected in this tree from that broad refresh to a narrow offline graph
repair: only the stale path-package dependency edges changed, while the
baseline registry package versions were retained. Therefore the Keynote/ops
`K` receipts still apply to their unchanged source; the standalone harness
receipts require a rerun against the corrected lock. Raw logs
remain outside the repository under
`/home/zhuhe/code/litchi-worktrees/.coord/keynote-lazy-entry/`.

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

The standalone harness lockfile was corrected offline because its tracked lock
predated the current workspace dependency graph. The corrected lock changes
only the stale path-package edges (`smallvec`, `soapberry-zip`, `siphasher`,
the `sha2 0.11.1` edge, and the missing `sha2` edge); it does not carry the
unrelated registry upgrades from the first refresh. The corrected lock is
included with this consumer migration and has sha256
`95aed16c860cd6082cd50a39aabf29779b89bdb004c8263674e042d84a007870`.
The earlier broad lock-generation log is external:
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

After the repair, focused and full IWA validation passed on overlay `I`:
clean source `8f747b7f0abe39ae9b5f098b0832b9b7967c2340` plus the dirty
one-line test-fixture edit described above. That exact overlay is committed in
`178ef2ff3c68b9a26ff6146a0f71eed3ef405194`:

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

The final assigned target tree was
`/home/zhuhe/code/litchi-worktrees/.coord/keynote-lazy-entry/target`, exactly
`6818426880` bytes by `du -sB1`. The environment fingerprints were:

```text
.cargo/config.toml                                      50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d
rust-toolchain.toml                                     e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310
Cargo.lock                                               6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da
tools/perf-baseline/unified-batch-closure/Cargo.lock     95aed16c860cd6082cd50a39aabf29779b89bdb004c8263674e042d84a007870
.coord/keynote-lazy-entry/env.sh                         42d98d07d010d128109f1a2b1d4cf58e78788ce576579fe71d7a97be1d0c008c
```
