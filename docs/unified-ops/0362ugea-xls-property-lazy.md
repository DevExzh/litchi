# 0362ugea — XLS property and worksheet-feature lazy entry

This slice closes the remaining XLS edit-entry debt from 0362uge. The
package metadata, OLE property-set, and worksheet-feature families now keep a
cheap clone of the immutable package owner at `edit_*()` entry. Their source
bytes, property-set parser/editor, typed metadata value, Workbook stream, and
BIFF globals are created only by the first accessor, verb, or commit.

The lazy shells retain the package's admitted policy. Property-set materialize
uses the CFB profile returned by `Package::edit_profile`, including the
source-backed CFB ceilings. A failed first materialization installs no partial
state, so a retry uses the same owner and policy. Metadata accessors are
fallible (`properties()` and `properties_mut()`), so malformed property
streams cannot be hidden behind an infallible reference. The XLS unified
metadata adapter, examples, and tests pass the new accessor result through.

Each family keeps its existing publication behavior: encrypted input is
refused before an edit is created; a source mutation after first materialize
is checked before property or worksheet-feature publication; exact no-ops keep
the source package and identity patch; changed unsigned commits republish
through the retained eager or source-backed policy; and dropping a cold edit
performs no source capture. Signed inputs can still be read and committed as
exact no-ops after a projection has been resolved, including metadata and
worksheet features, while a changed property, metadata, or feature commit is
refused before any publication. Signed no-op patches replay byte-for-byte;
explicit signature removal leaves the returned unsigned package editable. The
worksheet-feature edit uses private lazy stream state, so all feature,
drawing, note, hyperlink, validation, format and filter verbs materialize
through one gate.

## Scope and consumers

- `crates/litchi-xls/src/package/metadata.rs`: deferred typed properties and
  fallible property accessors and signed no-op/change policy.
- `crates/litchi-xls/src/package.rs`: crate-private signed-state projection
  used by the scoped publication guards.
- `crates/litchi-xls/src/workbook/package/property_set.rs`: deferred CFB
  snapshot/editor, retained CFB limits, source freshness check, and signed
  no-op/change policy.
- `crates/litchi-xls/src/sheet_features/edit.rs` and `objects.rs`: deferred
  Workbook stream/BIFF globals, all feature-object consumers, and typed
  signed-write refusal.
- `crates/litchi-ops/src/format/xls/metadata.rs` and
  `crates/litchi-ops/examples/ole_metadata_bench.rs`: accessor-result
  migration.
- `crates/litchi-xls/tests/xls_metadata.rs`: accessor-result migration,
  signed read/no-op/change/replay checks, and existing exact no-op, patch,
  inverse, and fidelity checks.
- `crates/litchi-xls/tests/sheet_features.rs`: signed read/no-op/change and
  replay checks.

Private unit checks assert that all three package edit constructors retain no
source payload at entry. Focused signed tests passed for metadata/property-set
and worksheet-feature read, no-op commit, changed-write refusal, and no-op
replay. The final validation packet below records the full and focused XLS,
Clippy, doctest, and affected consumer results.

The source fix was validated in the isolated capped target after the initial
source-only commit. The generated workspace lock is intentionally ignored and
kept outside Git at `/home/zhuhe/code/litchi-worktrees/uo2-xls-property-lazy/Cargo.lock`
(SHA-256 `6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`).
The signed metadata/property-set test passed 1/1 and the signed
worksheet-feature test passed 1/1. Their raw logs remain outside the checkout:
`signed-metadata-4f08-v3.log` (under `.coord/xls-property-lazy`) and
`signed-sheet-4f08-v2.log` (same directory). The broader focused XLS suite,
strict Clippy, and affected `litchi-ops` checks are recorded below. No native
qualification, benchmark, or performance claim is made here; no binary or
large evidence artifact belongs in this record.

The publication boundary now has a second, version-only original-source fence:
`Package::republish_bytes` checks it after candidate admission, so a render or
admission error cannot hide a concurrent source change. Property-set rendering
also fences after `PropertySetEdit::finish` and before returning either its
success or error, without another package-byte capture. Metadata republishing
uses this same boundary; worksheet-feature changed commits and replay already
route through it. The deterministic dynamic-source regressions cover both
property-set render and candidate publication refusal. The focused command was:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
cargo test -p litchi-xls --test package_owner --locked --offline property_set_ -- --nocapture
```

It passed 2/2; the external 592-byte log is
`/home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/property-fence-tests.log`
(SHA-256 `8fcb989656654c47369b24f62ad43b3c5f4ec57bd5370777e01bdf46f56bdfdd`).

## Final validation packet

The final source is commit `9fe06599d6be568218446b491312e1ce18a1e2e8`
(`895bb4905de2` is its publication-fence parent). Every final command below
was run with `git status --porcelain=v1` empty; the only workspace lock was the
ignored generated lock recorded above. The command environment was the
`.coord/xls-property-lazy/env.sh` file (493 bytes, SHA-256
`9de5d9faae447066b151f8414ac80f327506a9a5ab0dae72f03f079d141bd4da`):
Rust/Cargo 1.95.0 (`rustc 1.95.0 (59807616 2026-04-14)`, `cargo 1.95.0
(f2d3ce0bd 2026-03-21)`), `CARGO_BUILD_JOBS=1`,
`RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, debug info disabled, offline
locked resolution, and `LITCHI_MEM_MAX=20G`. The tracked root manifest
`Cargo.toml` is 10,672 bytes with SHA-256
`f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de`; the
tracked `.cargo/config.toml` is 470 bytes with SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`. The
tracked `rust-toolchain.toml` is 86 bytes with SHA-256
`e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310`; tracked
`clippy.toml` is 2,165 bytes with SHA-256
`f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c`. The
ignored `Cargo.lock` is 139,592 bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The final target was 3,343,185,454 bytes, below the 20,000,000,000-byte cap;
no target or executable is committed.

Final-source commands, all exit 0:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
cargo test -p litchi-xls --lib --locked --offline package::tests::republish_ -- --nocapture
```

2 passed, 0 failed, 0 ignored; 610-byte log
`republish-admission-tests.log`, SHA-256
`516423be57eddddfcf0e6973d1e65b669733e7538bd8e10c0952898a9b25f5cf`.
The stable control preserves the candidate CFB error; the hook control
mutates after the leading fence and receives typed `SourceChanged`.

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
export RUSTDOCFLAGS='-D warnings'
cargo test -p litchi-xls --all-features --doc --locked --offline -- --test-threads=1
```

The explicit XLS doctest gate passed 24 tests, 0 failed, 1 ignored (23 normal
doctests plus 1 compile-fail doctest); the 2,856-byte log
`xls-doctests-9fe.log` has SHA-256
`7754a17d16e0f2b370730671163bd22331d502bfaa14720aa2868e434cfeebd9`.

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
unset RUSTDOCFLAGS
cargo clippy -p litchi-xls --all-targets --all-features --locked --offline -- -D warnings
```

Strict all-target/all-feature XLS Clippy passed; the 168-byte
`xls-clippy-9fe.log` has SHA-256
`fef85535f7c1b27276d09a2761711f8cb49787342bdf0f2bdaa3dfa0bbdc281f`.

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
cargo test -p litchi-ops --all-features --test ole_metadata --test xls_features --locked --offline --no-fail-fast
```

The focused post-mapper consumers passed: `ole_metadata` 4/4 and
`xls_features` 7/7, with 0 ignored; the 1,297-byte
`ops-xls-consumers-9fe.log` has SHA-256
`d40be8705f4a7befa2e641a5f59f515c355fe90c89ace556eee103e7905ec424`.

The preceding publication-fence focused run on clean `895bb4905d` passed the
two `package_owner` regressions (2/2; 592-byte
`property-fence-tests.log`, SHA-256
`8fcb989656654c47369b24f62ad43b3c5f4ec57bd5370777e01bdf46f56bdfdd`) and
the metadata/worksheet suites (19 passed, 0 failed, 1 ignored; 1,652-byte
`metadata-sheet-final-fence.log`, SHA-256
`bce12e661e7ab31a3d01abd36db483e67c1873e0c565fc1d95a73e6a796f9ae3`).

Earlier full checks remain pinned to their exact pre-publication source and
are retained as corroborating evidence rather than relabeled as final-tip
results. On clean `04346195d46a230baec1e3f29b580d9250b1d0bf`,
`cargo test -p litchi-xls --lib --locked --offline --no-fail-fast` passed
1,323/1,323, 0 ignored (106,926-byte `xls-lib-0434.log`, SHA-256
`5ccb83b400b266346a4d603951f4a6e00ff491b3c46ae43a707e3cb8407b7a05`), and
`cargo test -p litchi-xls --all-targets --all-features --locked --offline
--no-fail-fast` passed 1,875 tests across 100 suites, 0 failed, 1 ignored
(169,905-byte `xls-all-targets-0434.log`, SHA-256
`00167a1cfa9433c6755bd7f9c4a7caac9bf760135924878e33737f2837a0721c`). Its
strict XLS Clippy passed (3,930-byte `xls-clippy-0434.log`, SHA-256
`8f07450ca8c96dbd43e1058776cd4d7ec5c741509769959b643bdfca15bdd97b`).
The earlier metadata consumer passed 4/4 (12,961-byte
`ops-ole-metadata-0434.log`, SHA-256
`736d8954911f3f784012037daa36def924f26030d53ab83fcb5b6f744172a312`);
the mapper-era XLS feature consumer passed 7/7 (725-byte
`ops-xls-features-c85.log`, SHA-256
`eafeb48d80fed30d402b121a6f12b5be5800162d705d90c651a60096580a8890`), and
the typed mapper unit passed 1/1 with the other all-feature targets filtered
(`ops-xls-native-c85.log`, 44,856 bytes, SHA-256
`2466fd7bf74d2003023cc523df210dd28da8ff1524f97699e7a3e8ab039b6aca`).

These are source/build and host-consumer checks only. No native qualification,
benchmark, or performance claim is made by this packet.

## Worksheet-feature fence successor

Commit `4dce73b23fb5f5673c19e379884f62ad4600b032` applies the same result
wrapper to worksheet-feature rendering and object-editor admission. A trailing
original-source fence now runs even when those stages return an error; the
successful path continues through the common republish fence. This is a source
successor to `9fe06599d6`; the earlier `9fe` receipts above are not relabeled.

On clean `4dce73b23f`, the focused command was:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
cargo test -p litchi-xls --test sheet_features --locked --offline --no-fail-fast
```

It passed 12/12 with 1 existing ignored evidence test; the external 917-byte
`sheet-features-4dce.log` has SHA-256
`65da5ad60352944e0c0179c75424b29ccd3e9e0f069fc2362e5d00de98df4510`.
Strict `cargo clippy -p litchi-xls --all-targets --all-features --locked
--offline -- -D warnings` also passed on the same clean source; its 376-byte
`xls-clippy-4dce.log` has SHA-256
`5e2cc8fa002efc7d30797ae25f313d99b2145c4f70d1ae4e3d1d256894a0f61c`.

Commit `3ae7b831406b01a801d5e75b171db0974243a13c` adds the causal controls
requested for that fence: the stable over-limit object-stream admission still
returns the concrete `StreamBytes` limit, while the source hook mutates after
the pre-admission observation and the same admission error is promoted to
typed `SourceChanged`. On clean `3ae7b83140`, the focused control command was:

```text
source /home/zhuhe/code/litchi-worktrees/.coord/xls-property-lazy/env.sh
cargo test -p litchi-xls --lib --locked --offline sheet_features::edit::tests::render_error_ -- --nocapture
```

Both controls passed (2/2; 0 ignored); the external 564-byte
`sheet-render-error-controls-3ae.log` has SHA-256
`e2b543362ec69ce280b4612cb67bccfac9d495e898fed92806479a9c977f71e8`.
The sheet suite was rerun on the same clean source (12 passed, 1 ignored;
1,025-byte `sheet-features-3ae.log`, SHA-256
`03fdd36671e24bfb64fd06e383cab69ce51747d9ce8aca2f032f78f80e74248a`), as
was strict all-target/all-feature Clippy (168-byte `xls-clippy-3ae.log`,
SHA-256 `a9833e5e165bfbd444f5b4b1f2aa43eaa60c69c51e41fdf5b5ecc8182d4286aa`).
