# 0362ugjab — PPT lazy-native preparation and bounded compile return

This return corrects the preparation gap in 0362ugja. The first baseline
release-build attempt copied the three probe sources and the temporary
`src/lib.rs` registration, but did not materialize the reviewed PPT fixture
closure in the historical checkout. The compiler therefore failed at the
`include_bytes!` boundary. The corrected preparer now copies and hashes every
tracked external `include_bytes!` input needed by the standalone
`perf-baseline` library and PPT probe into the external baseline endpoint,
records each source and destination identity, and preserves the unchanged
historical production tree. These files are never added to the historical
Git commit or to a production overlay.

## Source-only correction

The new external-only preparer is
`tools/native_qualification/prepare_ppt_lazy_native_inputs.py`. It requires
the exact pre-lazy baseline ref/tree and standalone lock, a clean candidate
worktree, ordinary tracked candidate probe and fixture files, and a new
receipt path outside both worktrees. It materializes exactly these temporary
paths in the baseline endpoint:

```text
tools/perf-baseline/src/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs
tools/perf-baseline/src/lib.rs
test-data/unified-ops/ppt-lazy-native-fixture.ppt
test-data/poi/test-data/slideshow/ppt_with_embeded.ppt
test-data/unified-ops/0294/video.ppt
test-data/poi/test-data/slideshow/bug61881.ppt
test-data/poi/test-data/spreadsheet/54016.xls
test-data/rtf/watermark.rtf
```

The first three are copied byte-for-byte; `lib.rs` receives only the guarded
`pub mod ppt_lazy_native;` insertion; and all six fixtures are copied
byte-for-byte from the candidate's ordinary tracked files, except for the
five compile fixtures that are already in the pinned baseline Git tree: a
sparse checkout may omit their worktree files, so the preparer verifies their
HEAD blob identities and restores them with `checkout-index
--ignore-skip-worktree-bits`. It rejects an existing untracked or ignored file
at any reviewed fixture path. The two
`include_str!("lib.rs")` uses remain self-contained source reads and require
no external fixture. The receipt records the candidate and baseline ref/tree,
lock, every source and destination identity, the sparse pre-state,
registration before/after hashes, the complete path allowlist, and the
baseline production-tree digest. The pinned baseline ref/tree and permitted
sparse omission state are checked before any materialization. It claims
preparation identity only. The runtime runner's build-receipt validator now
requires this preparation receipt and the complete fixture binding before
accepting a correctness report.

The fixed compile-fixture identities are:

| path | bytes | SHA-256 |
| --- | ---: | --- |
| `test-data/poi/test-data/slideshow/ppt_with_embeded.ppt` | 180,736 | `afa1dd5adc36765e6e40725a133b0330e5ad1febd1506ca93ceae8c5c5400b53` |
| `test-data/unified-ops/0294/video.ppt` | 43,520 | `f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a` |
| `test-data/poi/test-data/slideshow/bug61881.ppt` | 74,752 | `ba2bf551e87f5fd7667dec52ac19005396bdf9af00f0006f9a247a9fede2e8ee` |
| `test-data/poi/test-data/spreadsheet/54016.xls` | 984,576 | `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a` |
| `test-data/rtf/watermark.rtf` | 69,471 | `48d62dcd959e737b06ebb8255780bcaaf1e88056ff9c3d5a21d3ff5cd3ddf9cb` |

The regression is covered by:

```text
PYTHONDONTWRITEBYTECODE=1 PYTHONPATH=tools \
  python3 -m unittest -v \
  tools/test_prepare_ppt_lazy_native_inputs.py \
  tools/test_ppt_lazy_native_qualification.py
```

Both preparation tests and all 13 existing PPT protocol tests passed. The
reviewed fixture is 14,848 bytes with SHA-256
`02fa44ca76903943e80a3a2da1f9dbd326c7a4a047540c3aa9cf211e34ef617c`.
The focused suite now has 18 tests, including a temporary sparse worktree,
wrong-baseline no-mutation, and pre-existing-untracked-fixture rejection.

The actual CLI was rerun against a fresh sparse checkout after these guards:

```text
receipt: /home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/preparation-sparse-cli.json
receipt SHA-256: 64051f07b614270cfb1769ecbc9752be6a16c7e6efb63daf79c0a77fa264c3fe
baseline: 8779d059add6c90c74552801372c44f770db9f0f / 27e0b56e2f8e3ab11b00c119604b5636f358dfe2
candidate: fb262b0bbf5160534ca5e87c63727df823a794bc / cb0965bc806b3f7c7a38f3bc340499039bf7e726
candidate lock: 6dbfc0bcc64cdf8e625a6562f2d5d11c1e98ad7cb6edda299b9adf2a36a70e39
```

The receipt binds ten reviewed materialization paths and all five compile
fixtures; no build or probe was run by the preparer.

## Preserved failed attempts

These are diagnostic receipts only; neither was relabelled as a successful
build or correctness run. Both commands were actual release builds, not
`cargo check`, and neither executed a probe:

| endpoint | source ref/tree | command | exit | log (bytes; SHA-256) |
| --- | --- | --- | ---: | --- |
| baseline | `8779d059add6c90c74552801372c44f770db9f0f` / `27e0b56e2f8e3ab11b00c119604b5636f358dfe2` | `cargo build --release --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin ppt_lazy_native` | 101 | `/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/baseline-8779-normal-build.log` (10,936; `db4d630ce70d67ab589b7c6cf4818bacbd4df5ef8904c532ffe6b003e3b77ceb`) |
| old candidate | `4fbfe1205a2a9c3842cd2a439323cb98d41d7265` / `4d7de85cb49f42189b703c6cc0267997b01bfffa` | `cargo build --release --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin ppt_lazy_native` | 101 | `/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/candidate-4fb-normal-build.log` (9,995; `281ca113aeb9231dfab01f846d06d4865cfbea9a185e5d342b721801eb224a0c`) |

The baseline log ends at the missing default fixture. The old candidate log
ends at `litchi-doc/src/package/model.rs:730`, where the pre-merge source
matched a non-exhaustive `litchi_vba::Error` without a wildcard. The corrected
candidate source is the explicitly selected coordinator staging source, not
the old 4fb checkout; its final ref, source archive/root identity, lock and
build receipt must be filled together for the bounded rebuild. No dependency
graph is regenerated implicitly.

The first corrected baseline retry used the default fixture but still lacked
the rest of the tracked compile closure. It was therefore preserved as a
failed build, not relabelled as a qualification result:

| endpoint | source ref/tree | command | exit | log (bytes; SHA-256) |
| --- | --- | --- | ---: | --- |
| baseline retry r2 | `8779d059add6c90c74552801372c44f770db9f0f` / `27e0b56e2f8e3ab11b00c119604b5636f358dfe2` | `cargo build --release --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin ppt_lazy_native` | 101 | `/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/baseline-8779-r2-normal-build.log` (12,637; `d3727cdaed9fc186df484f5434b5fb72c9455cf34a1a93db4094625e187942b4`) |

Its diagnostics identified the five additional tracked inputs now bound by
the preparer: `ppt_with_embeded.ppt`, `video.ppt`, `bug61881.ppt`, `54016.xls`
and `watermark.rtf`.

## Bounded rebuild boundary

The corrected pair is built only under the repository memory-capped Cargo
shim, with one build job, offline and locked resolution, incremental
compilation disabled, release profile, and separate external targets below
20,000,000,000 bytes. The historical baseline retains lock
`4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` (57,158
bytes). The candidate source is the current reviewed coordinator staging
tip selected for the rebuild; its lock identity is recorded explicitly in the
preparation/build receipts. Source archive/root, toolchain, Cargo config,
target paths and binary identities are all external receipt inputs.

The only permitted next execution is one normal release build per endpoint
followed by the probe's `--mode verify` correctness invocation through the
existing fail-closed runner. Allocator builds, calibration, timing, ABBA and
performance disposition remain pending and are not claimed here.

## Final probe-fix build and correctness preflight

The probe correction was committed as `416f57cd7a` in the reviewed source
packet. It changes absent optional patch endpoints in a no-op proof from the
boolean `false` to JSON `null`; the validator requires the distinction and
the source probe hash below binds the exact corrected file. The final
candidate adapter was built at temporary commit
`5067a9eac047ce40fc714c16af7bcfa100a229af` (tree
`cb99f0c28121626e1e110ed803674ee2021c5c73`), which is the local-lock
migration `fb262b0bbf5160534ca5e87c63727df823a794bc` plus that probe fix.
The final preparation receipt is retained outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/preparation-5067a9eac0.json`
(12,580 bytes; SHA-256
`0bbb366e920949da8f212b7a0bb2ad1c4499054cd810f5b64605078f58e573a9`).
It binds the baseline ref/tree and the candidate ref/tree above, the
baseline lock (57,158 bytes; SHA-256
`4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184`), the
candidate lock (57,206 bytes; SHA-256
`6dbfc0bcc64cdf8e625a6562f2d5d11c1e98ad7cb6edda299b9adf2a36a70e39`), the
five compile fixtures, the sparse pre-state, and the complete overlay.

The three probe source identities in that receipt are:

```text
tools/perf-baseline/src/ppt_lazy_native.rs          24,586  68d77fad390dc681ed248b2348b26f319d2e646d1f0b006796ecd89c49cae62e
tools/perf-baseline/src/bin/ppt_lazy_native.rs        146  04df53920b036911e9002df152b9e07b1e0392ac6a1a545b285807108d808d12
tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs  203  661e94115adc418de1c25ee6dd5481e4399e43eee26e011a49d1fb9d1ce56201
```

With the repository memory-capped Cargo shim, `LITCHI_MEM_MAX=20G`, one
build job, `CARGO_INCREMENTAL=0`, offline and locked resolution, both
endpoints completed the recorded release build command:

```text
cargo build --release --offline --locked \
  --manifest-path tools/perf-baseline/Cargo.toml --bin ppt_lazy_native
```

The build used rustc 1.95.0 (commit
`59807616e1fa2540724bfbac14d7976d7e4a386b`, LLVM 22.1.2), and separate
external targets below the 20-GB ceiling. The retained identities are:

| endpoint | binary (bytes; SHA-256) | build log (bytes; SHA-256) | target bytes |
| --- | --- | --- | ---: |
| baseline | 12,130,968; `ad031ccabb073b3a719f0985088f43ee2ab9a295c9e52594a70841d8e2090a1f` | `/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/baseline-8779-r4-probe-fix-build.log` (579; `8c584eefe62f7a91620aa6167fdbe62e18abee0e26a9200049ae6a1a3ffa3e20`) | 1,314,509,577 |
| candidate | 12,289,336; `a62114de50bc25eaf4ba5682c632940efe3a97a1d7140dcb323105c4b1981a61` | `/home/zhuhe/code/litchi-worktrees/.coord/ppt-lazy-native-compile-20261024/logs/candidate-5067-r2-probe-fix-build.log` (224; `bd6898ccea415c2a13310fea6fe119868092b866a1d73db9e9e8190fd5a39d78`) | 1,317,080,259 |

The normal binaries were then invoked separately with exactly
`--mode verify`. Both exited 0. Their 12,094-byte reports are retained at
`verify-baseline-8779-r4-probe-fix.json` and
`verify-candidate-5067-r2-probe-fix.json` beside the targets; both have SHA-256
`2daa5587982ac38c4a4068adc03a9c91798a4c98ce95b050e541069ed276d07c`.
The fail-closed Python validator accepted both reports: each proves all 28
routes over the four reviewed fixtures, and the validated correctness
structures are equal. This is a correctness-only preflight. It makes no
allocator, calibration, timing, ABBA, or performance-acceptance claim; those
remain pending a separate quiet-host authorization.
