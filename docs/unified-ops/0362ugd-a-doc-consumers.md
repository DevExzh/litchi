# 0362ugd-a — DOC lazy-entry consumers and profiler state

- **Status:** source follow-up on `uo2/doc-lazy-entry`; scoped validation is
  complete and the branch is pending coordinator integration.
- **Date:** 2026-10-24.
- **Input:** [0362ugd](0362ugd-doc-lazy-entry.md), implementation
  `663c1d905388964cfd6a2b516eefcf5efffb273d`.
- **Scope:** consumers and diagnostics affected by the fallible lazy DOC
  primary edit entry. No new editing family or capability was added.

## Consumer migration

The DOC facade now propagates the fallible `properties_mut()` getter and the
fallible `source_bytes()` getter. The coalesced body-run snapshot helpers take
`&mut self` where the getter can materialize the native body owner, and their
errors remain DOC transaction errors rather than being discarded. Ordinary
native edit storage and all direct DOC facade/tool signatures use
`litchi_doc::Edit<'static>`; PPT and XLS metadata consumers remain with their
format owners.

## Observer-state correction

The ordinary `Edit` carries a zero-sized default deferred-open state. A
profiled edit carries its borrowed observer state as a separate generic state,
so enabling `performance-diagnostics` does not make an ordinary edit lose the
`Send`/`Sync` auto traits it had before the lazy-entry change. The profiled
state continues to accept caller-scoped, non-`Send` callbacks. The package
profiling helper explicitly performs the first body access while its recorder
is live, then converts to the ordinary edit before the recorder is finished;
the open phase therefore measures real lazy materialization instead of zero
events. Failed profiled opens retain both observers for a later attempt, and
the profiled strict/fallback path receives the selected protection policy just
like the ordinary path. The DOC transaction coordinator propagates the
fallible coalesced snapshot charge without swallowing its error.

## Checks

The full tracked consumer audit used `git grep HEAD -- crates tools` and
separated the DOC-owned sites from the PPT/XLS sites assigned to their
workers. Direct `rustfmt --config skip_children=true --edition 2024` on all
changed Rust files and `git diff --check` passed.

The final scoped Cargo receipt is below. The DOC tests and Clippy checks ran at
the clean `eae0d57a146e56062e4306dc7e9f924504b10dc4` source commit. The later
`57dc30d4148621fc02eeaf05c9584799e0490f7d` commit changes only the format-gated
ops test helpers, and `251f33f1ce0da0bd9698ab41a2080a57cb7f1ec3` changes only
the two profiler-tool consumers; therefore the DOC paths exercised by those
checks are byte-identical in the final source. The ops check ran at clean
`57dc30d4148621fc02eeaf05c9584799e0490f7d`.

The two profiler-tool checks ran with `HEAD=57dc30d4148621fc02eeaf05c9584799e0490f7d`
and a dirty two-file overlay. That overlay is exactly the committed
`57dc..251f33f1ce` patch (`tools/perf-baseline/src/lib.rs` and
`tools/perf-cfb-profile/src/lib.rs`), SHA-256
`fdafc4ea35047a51b93ed9c44918966d66bcb126a02f8cf844aba77edd7273da`, 1,143
bytes. The resulting final source is clean at `251f33f1ce0da0bd9698ab41a2080a57cb7f1ec3`;
no uncommitted fix is omitted from that commit.

All commands used `(deleted-coord-artifact)`
with `--offline --locked`, `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`,
`CARGO_INCREMENTAL=0`, dev `debug=0`, and `LITCHI_MEM_MAX=20G`. The pinned
toolchain was `rustc 1.95.0 (59807616 2026-04-14)` and `cargo 1.95.0
(f2d3ce0bd 2026-03-21)`. The external target was
`(deleted-coord-artifact)`, measured at
1,605,542,020 bytes after validation; no target or raw log is committed.

| Command/source | Result | External raw log (bytes; SHA-256) |
|---|---|---|
| `cargo test --offline --locked -p litchi-doc --tests` at `eae0d57a14` | library: 1,211 passed, 0 failed, 2 ignored; every integration target passed | `(deleted-coord-artifact)/litchi-doc-tests-eae0d57a14-default.log` (129,385; `d0cc1792b9fc2a0cee407017a36d69e312700ddfb287afb51691614cb6cd7b27`) |
| `cargo test --offline --locked -p litchi-doc --lib --features performance-diagnostics` at `eae0d57a14` | 1,223 passed, 0 failed, 2 ignored | `(deleted-coord-artifact)/litchi-doc-tests-eae0d57a14-profiled.log` (102,246; `005e71d8ab1fa5a78e408f10d6b4bd062e26bad10fbc3dfcc8a8fab87b420c54`) |
| `cargo clippy --offline --locked -p litchi-doc --all-targets -- -D warnings` at `eae0d57a14` | passed | `(deleted-coord-artifact)/litchi-doc-clippy-eae0d57a14-default.log` (165; `7b35427d8b00929f227258b6f90a165a48b101f8bb7dec434d8f926f6f5dee94`) |
| `cargo clippy --offline --locked -p litchi-doc --all-targets --features performance-diagnostics -- -D warnings` at `eae0d57a14` | passed | `(deleted-coord-artifact)/litchi-doc-clippy-eae0d57a14-profiled.log` (165; `0c33b2880257d00a3cb0dc2b4baabd81670ef27beb7c382ab7eafd289c605426`) |
| `cargo check --offline --locked -p litchi-ops --all-targets --features doc` at `57dc30d414` | passed | `(deleted-coord-artifact)/litchi-ops-check-57dc30d414-doc-all-targets.log` (165; `734d1c46725ddf7d2081bc9b26165aa0925c934bdc84197a3b65507aa7c47770`) |
| `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets` on the recorded overlay | passed, no warnings | `(deleted-coord-artifact)/perf-baseline-check-final.log` (229; `a7f88396bad1948be94d3cf92fd9318192f90e6ae407226580906bff52aaf3a2`) |
| `cargo check --offline --locked --manifest-path tools/perf-cfb-profile/Cargo.toml --all-targets` on the recorded overlay | passed, no warnings | `(deleted-coord-artifact)/perf-cfb-profile-check-final2.log` (196; `1d7702ec8dbf78ec55407f61d306f3aefbe8c41ee5439738c33eb32a8c7ba7a3`) |
