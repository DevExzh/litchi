# 0362umc — OPC retained source artifacts release the read-ahead window

This receipt covers the OPC acceptance repair on top of
`410bd99f00fdfb4161a94215710340f16ea536bb`.  It is intentionally independent
of the coordinator's `PROGRESS.md`; the coordinator owns that file and the
merge.

## Failures and cause

Frozen acceptance leaf 0033 reported 1,603 retained bytes after the last
`SourceBackedPackage` owner had dropped, while the required assertion was zero.
The retained `SourceArtifact` shared the indexed archive, whose
`SourceReader` also held a strong `Arc<ArchiveReadAhead>`.  The artifact's
O(1) archive/catalog handle therefore pinned the bounded read-ahead allocation.
Leaf 0036 reported four OPC doctests that used `crate::SourceCacheLimits` and
did not import the public `SourceLimits`/`SourceCacheLimits` names.

## Repair

`SourceReader` now keeps `Weak<ArchiveReadAhead>`.  The live package still owns
the strong `Arc`; only the configured read-ahead branch attempts one checked
weak upgrade per positional read.  Exact-policy readers keep
`read_ahead: None` and go directly through the checked exact source path.  If a
configured reader's package owner is gone, it also falls back to
`read_source_at_with_context`, retaining the normal source-version, I/O and
cancellation fences.  Thus an artifact keeps the immutable archive index and
admitted catalog O(1) without retaining the bounded window, while live package
clones continue to share it until the final clone drops.  No unchecked native
getter, eager scan, reopen, catalog copy or public snapshot route was added.

The four affected OPC doctests now import `SourceLimits` and
`SourceCacheLimits` from `litchi_opc` and use those public names.  No aliases or
deprecated compatibility names were introduced.

The new integration controls in `source_read_ahead.rs` cover the exact
invariants behind the repair:

* `retained_source_artifact_streams_after_its_last_package_owner_drops`
  exercises `restore_streamed_artifact_to_stream`, including the recipient's
  `admit_source_artifact` path, after all source package owners are gone.
* `a_live_package_clone_keeps_the_read_ahead_window_until_the_last_clone_drops`
  proves the live strong owner remains shared and budgeted through the first
  drop.
* `dropped_artifact_freshness_precedes_restore_and_fingerprint` proves both
  post-drop fingerprint and restore return `SourceChanged` before output when
  the retained source changes.

The original retention assertion remains unchanged; it was not weakened.

## Validation

All commands used the isolated target
`/home/zhuhe/code/litchi-worktrees/.coord/final-opc-retention/target`,
`/home/zhuhe/code/litchi-worktrees/.coord/bin` first on `PATH`,
`env -u LITCHI_CAPPED -u LITCHI_BLESS`, `LITCHI_MEM_MAX=20G`, one Cargo job,
one test thread, incremental compilation disabled, debug profiles disabled,
offline locked resolution, and the isolated `TMPDIR`.

The final hydrated runs were:

* `cargo test --offline --locked --target-dir <target> -p litchi-opc --all-features --lib --no-fail-fast` — 714 passed, 0 failed.  Log: `opc-native-lib-hydrated.log`, 68,644 bytes, SHA-256 `1acdfc3c1416a9981bd2990f40d7d9b58953e627f09ddfa293336d0d259e4275`.
* `cargo test --offline --locked --target-dir <target> -p litchi-opc --all-features --tests --no-fail-fast` — 1,207 passed, 1 ignored, 0 failed.  The ignored test is the existing `independent_zip64_source_metadata_and_small_xml_edit_preserve_large_member`, which requires `LITCHI_0415_PYTHON_ZIP`; the source-read-ahead target was 19/19.  Log: `opc-native-tests-hydrated-final.log`, 118,275 bytes, SHA-256 `a5bc07b27f1a9463f5198bb9e0cf58ba0234e0c4d3be4619ed33d83289b8b842`.
* `cargo test --offline --locked --target-dir <target> -p litchi-opc --all-features --doc --no-fail-fast` — 10 ordinary doctests and 1 expected compile-fail doctest passed.  Log: `opc-doctests-hydrated.log`, 1,580 bytes, SHA-256 `7519fabbb4b46cb2db0c5d99aec315213fec830bc7857a786fd97531a69dc68e`.
* `cargo clippy --offline --locked --target-dir <target> -p litchi-opc --all-targets --all-features -- -D warnings` — clean.  Log: `opc-clippy-all-targets-final.log`, 169 bytes, SHA-256 `ef9d2fc61386373dde2a9a78f9a15c73d2329f2be890c1fd4dff9192dc997669`.
* `cargo test --offline --locked --target-dir <target> -p litchi-opc --all-features --test source_read_ahead --no-fail-fast` — 19/19 passed with the three controls.  Log: `opc-source-read-ahead-controls.log`, 2,001 bytes, SHA-256 `1c46e8ca8e287db56b503037c0e9e3aec505f201210acbd0f18da408cf5a1cfe`.
* `cargo fmt --all -- --check` and `git diff --check` — clean.  The format log is empty (0 bytes, SHA-256 `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`).  The later module-introduction wording correction is documentation-only.  A cheap final check ran `git show cdbee64a36fc89d11f19b60f7bfe3f9ff33021e5:crates/litchi-opc/src/source_backed/read_ahead.rs | rustfmt --check --edition 2021` and passed; its 55-byte log is `rustfmt-reviewed-source-cdbee.log`, SHA-256 `2627e6a1f85e7c9f99f1000c51598b0637a7018d0785772e0d3daf62f7b94a86`.

The checkout initially had sparse patterns excluding tracked `test-data/` and
the forced-ZIP64 corpus, which caused setup-only fixture errors in the first
native/Clippy attempts.  The worktree was hydrated with `/test-data/**` and
`/docs/performance/results/change-0416/corpus/**`; `3rdparty/` stayed excluded.
The successful runs above therefore use tracked fixtures and do not treat the
sparse checkout failures as exceptions.

The final target occupied 874,825,776 bytes.  The lockfile copied into this
worktree came from
`/home/zhuhe/code/litchi-worktrees/.coord/odp-final-f6fe-source-20261024/Cargo.lock`;
it is 139,591 bytes with SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.  The
same size and hash were independently observed at
`/home/zhuhe/code/litchi-worktrees/uo2-budget-validation/Cargo.lock`; that is a
comparison, not the copy origin.

## Required performance qualification

The source-read path changed, so this receipt does not claim a timing win.
After the gate build/load subsides, run an ABBA-quiet comparison of the exact
baseline endpoint `410bd99f00fdfb4161a94215710340f16ea536bb` and exact candidate
endpoint `cdbee64a36fc89d11f19b60f7bfe3f9ff33021e5` with the existing
`tools/perf-baseline/src/docx_managed_read_ahead.rs` production harness.  Keep
the locked dependencies, toolchain, source corpus, build flags and machine
matched; the source revision is the respective endpoint SHA, not one shared
SHA.  Bind each before/after binary and output to its real endpoint SHA.  Use
the `docx-managed-read-ahead` selector with the existing 4 KiB
`forward-start` policy and the exact-policy control, including a zero-
artificial-delay route; keep setup and artifact-retention memory controls
outside the timing scope and report them separately.  The matched commands
must use the harness's `--policy`, `--window-bytes`, `--max-range`,
`--delay-us`, `--samples`, `--warmup`, `--source-revision` (the corresponding
endpoint SHA) and `--output` arguments, with before/after binaries built from
the same lock and flags.  Exact-policy controls must show no weak upgrade;
the qualification must measure the configured live path's one weak upgrade
per positional read explicitly rather than hiding it behind a generic
source-read result.  The expected memory result is no retained window after
package drop.
