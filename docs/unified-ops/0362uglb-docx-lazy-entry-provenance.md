# 0362uglb — DOCX lazy-entry validation provenance

This documentation-only successor disambiguates the validation in
`0362ugla-docx-lazy-entry-validation.md`.  No Cargo command was rerun for this
record.  The received leaf includes documentation through `3a6c56958c`; its final
source tree (before the documentation-only records) is `0cf1a17174fffe46f970b0399ef39bf803cdebce`.
The behavior-bearing validation pins are kept distinct below:

* `e5bc620a13681d181e73c6e367f323e00af902fd` is the DOCX and facade test,
  doctest and DOCX Clippy source pin.
* `609523b702e6ac2de20e725debd96ffbab5ad9f8` is the signed empty-property
  correction and the final litchi-ops test, focused test and litchi-ops
  Clippy source pin.
* `472d21297b9df437f1dd30a270c1faaabd75fce5` is the committed standalone
  publication-consumer migration used by the litchi-ops fuzz and
  perf-baseline checks.
* `0cf1a17174fffe46f970b0399ef39bf803cdebce` is the clean source pin for the
  final format checks.  The remaining `0372b48491` change is documentation.

The only captured command with a dirty source overlay was the DOCX fuzz check:
it ran at base `609523b702` with the two publication-consumer edits in
`crates/litchi-docx/fuzz/fuzz_targets/tail_append.rs` and
`source_backed_tail_append_stream.rs`; those edits were committed immediately
afterward as `472d21297b`.  The 5,505-byte committed two-file overlay has SHA-256
`a2aeeb774ca607ad3f3d1a2ae5823eb23209ff48ae4f59e93f6e2aafdd793284`.
Every other command below ran with a clean
worktree at the stated pin.

## Toolchain and policy identities

The coordinator environment was sourced from
`/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/env.sh`, SHA-256
`3b2855875112ff5e830c55e82e22baec0e5711252023099d83a06a6954539008`.  It set
`LITCHI_MEM_MAX=20G`, `CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, disabled
incremental/debug artifacts, enabled Cargo offline mode, and selected
`/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/target` and its
`tmp` directory.

The source/config identities were unchanged at all four validation pins:

| file | SHA-256 | relevant identity |
| --- | --- | --- |
| `Cargo.toml` | `f5ec007892023b38f1deaef1ce20714db58b58448a3bea088be3b847b84570de` | workspace resolver 3, `crates/*`, edition 2024, workspace `clippy::all` denied with the documented Rust 1.95 compatibility allows |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` | channel 1.95.0, clippy and rustfmt, minimal profile |
| `.cargo/config.toml` | `50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d` | `lint = "clippy --workspace --all-features --lib --no-deps -- -D warnings"` |
| `clippy.toml` | `f11517bc5872d96aab42ba6519db4e78f2fdcd44512ac70a1b45adbd9929b50c` | 2,165-byte tracked workspace Clippy configuration |
| `tools/perf-baseline/Cargo.lock` | `4def018e9c1d5f409b8c60d6cc0b8f114c486ab3f8fed2c7a4e9c7baf21c5184` | tracked standalone perf-baseline lock |

The workspace `Cargo.lock` used by the `--locked` workspace commands was the
ignored generated lock with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
The two standalone fuzz commands intentionally did not use `--locked`; their
generated ignored lock identities were DOCX
`8aae6bb4f94e0582c0bbd48e5d0776b2ab5110e131aa0cc3346f3c170e8e7928` and
litchi-ops
`10708241205e4a90c0120b3f41a77e7548715761801d66da2d55911e59e9ab8a`.

The executable identities were `rustc 1.95.0 (59807616e 2026-04-14)`,
`cargo 1.95.0 (f2d3ce0bd 2026-03-21)`, `clippy 0.1.95 (59807616e1
2026-04-14)`, and `rustfmt 1.9.0-stable (59807616e1 2026-04-14)`.

## Exact commands and source overlays

Each command was run after `source
/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/env.sh`; the command
column preserves the Cargo flags that controlled the gate.

| validation | exact command | source / overlay | lock |
| --- | --- | --- | --- |
| DOCX all targets | `cargo test --offline --locked --no-fail-fast -p litchi-docx --all-features --all-targets` | `e5bc620a13`, clean | workspace generated lock |
| DOCX doctests | `cargo test --offline --locked -p litchi-docx --all-features --doc` | `e5bc620a13`, clean | workspace generated lock |
| facade all targets | `cargo test --offline --locked --no-fail-fast -p litchi --features docx --all-targets` | `e5bc620a13`, clean | workspace generated lock |
| litchi-ops all targets | `cargo test --offline --locked --no-fail-fast -p litchi-ops --features docx --all-targets` | `609523b702`, clean | workspace generated lock |
| signed regression | `cargo test --offline --locked -p litchi-ops --features docx --test docx_unified_write empty_assignments_keep_security_policy_and_signed_noops_are_retained -- --exact --nocapture` | `609523b702`, clean | workspace generated lock |
| DOCX strict Clippy | `cargo clippy --offline --locked -p litchi-docx --all-features --all-targets -- -D warnings` | `e5bc620a13`, clean | workspace generated lock |
| litchi-ops strict Clippy | `cargo clippy --offline --locked -p litchi-ops --features docx --all-targets -- -D warnings` | `609523b702`, clean | workspace generated lock |
| DOCX fuzz manifest | `cargo check --offline --manifest-path crates/litchi-docx/fuzz/Cargo.toml --all-targets` | `609523b702`, dirty two-file overlay later committed as `472d21297b` | generated DOCX fuzz lock |
| litchi-ops fuzz manifest | `cargo check --offline --manifest-path crates/litchi-ops/fuzz/Cargo.toml --all-targets` | `472d21297b`, clean | generated litchi-ops fuzz lock |
| perf-baseline consumers | `cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin docx_plain_paragraph_tail_append --bin docx_bounded_tail_append_compare --bin docx_replayable_tail_append` | `472d21297b`, clean | tracked perf lock |
| DOCX formatting | `cargo fmt --manifest-path crates/litchi-docx/Cargo.toml -- --check` | `0cf1a17174`, clean | none |
| DOCX fuzz formatting | `cargo fmt --manifest-path crates/litchi-docx/fuzz/Cargo.toml -- --check` | `0cf1a17174`, clean | none |
| litchi-ops formatting | `cargo fmt --manifest-path crates/litchi-ops/Cargo.toml -- --check` | `0cf1a17174`, clean | none |
| litchi-ops fuzz formatting | `cargo fmt --manifest-path crates/litchi-ops/fuzz/Cargo.toml -- --check` | `0cf1a17174`, clean | none |
| whitespace check | `git diff --check` | `0cf1a17174`, clean | none |

The package test totals are 2355 DOCX all-target tests, 739 litchi-ops tests
and 23 facade tests.  The DOCX doctest total is 170 cases: 168 ordinary
doctests plus 2 compile-fail doctests; 31 ordinary doctests were ignored.  The
all-target totals above count only executed passing tests; those three runs
had zero ignored tests.  All commands
returned exit code 0.

The target directory measured 4,803,301,769 bytes.  The applicable decimal
cap is 20,000,000,000 bytes; the `20G` environment setting is a process-scope memory
cap and is not the target-size qualification.  The perf-baseline command was
compile-only, so this record makes no native timing or ABBA qualification
claim.

## Captured command logs

These are the captured terminal logs, not reduced summaries; they remain
outside the checkout and were not committed.

| command | external log | bytes | SHA-256 |
| --- | --- | ---: | --- |
| DOCX all targets | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-all-target-tests-e5bc620a13.log` | 222427 | `7d62612f1383859166970cff183a9542b5a7c0775779f74eaadec51c0563f3fe` |
| DOCX doctests | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-doctests-e5bc620a13.log` | 23348 | `31bb84388896c2e55f831d44d04c5590f230165f71b5e663a8ab27b128feba5c` |
| facade all targets | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-docx-tests-e5bc620a13.log` | 11996 | `e05bff5fed9a5d0af1b1f06a4b67ca03e766d2edae0c8e1409625aac2738e8ec` |
| litchi-ops all targets | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-docx-tests-609523b702.log` | 117399 | `d54a30657fbcc0da12bd7edd634783acf5775bfa124a23bcbde79a8af316cd3b` |
| signed regression | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-signed-empty-correction.log` | 515 | `6d57b6e7b372cd8a8cfa13edac88637ceb2cc616a2c7c257b5c4420c95d9ae21` |
| DOCX strict Clippy | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-clippy-e5bc620a13.log` | 168 | `519c93b7d35693e5e7c13590d69d3ce59114fb765c0172aa57a8444cb4fb2100` |
| litchi-ops strict Clippy | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-docx-clippy-609523b702.log` | 166 | `5174c1853f86a3fe8335d9c9d30c5b5eed97a597cfd4e0a7d4d20a2ba40da819` |
| DOCX fuzz manifest | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-fuzz-check-final2-609523b702.log` | 177 | `60e0d59bcb018462f6efa1882ad7622c9e01df31cf984440295cf1f877d7827a` |
| litchi-ops fuzz manifest | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/litchi-ops-fuzz-check-472d21297b.log` | 951 | `ee209e1ceeddedbb6534a614ddb8da221234e5c14320bb056f161aef9e0888af` |
| perf-baseline consumers | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/receipts/docx-perf-baseline-check-472d21297b.log` | 8549 | `256d140ede2bd24db2c2536bdfdd7046436cc8152ce4ad74a9832a3cccf71268` |
