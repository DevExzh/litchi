# 0362ugjab — PPT lazy-native preparation and bounded compile return

This return corrects the preparation gap in 0362ugja. The first baseline
release-build attempt copied the three probe sources and the temporary
`src/lib.rs` registration, but did not materialize the reviewed default PPT
fixture in the historical checkout. The compiler therefore failed at the
`include_bytes!` boundary. The corrected preparer now copies and hashes that
fixture into the external baseline endpoint, records its source and
destination identities, and preserves the unchanged historical production
tree. The fixture is never added to the historical Git commit or to a
production overlay.

## Source-only correction

The new external-only preparer is
`tools/native_qualification/prepare_ppt_lazy_native_inputs.py`. It requires
the exact pre-lazy baseline ref/tree and standalone lock, a clean candidate
worktree, ordinary tracked candidate probe and fixture files, and a new
receipt path outside both worktrees. It materializes exactly these five
temporary paths in the baseline endpoint:

```text
tools/perf-baseline/src/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native.rs
tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs
tools/perf-baseline/src/lib.rs
test-data/unified-ops/ppt-lazy-native-fixture.ppt
```

The first three are copied byte-for-byte; `lib.rs` receives only the guarded
`pub mod ppt_lazy_native;` insertion; and the fixture is copied byte-for-byte
from the candidate's ordinary tracked file. The receipt records the
candidate and baseline ref/tree, lock, source and destination identities,
registration before/after hashes, the five-path allowlist, and the baseline
production-tree digest. It claims preparation identity only. The runtime
runner's build-receipt validator now requires this preparation receipt and
the exact fixture binding before accepting a correctness report.

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
