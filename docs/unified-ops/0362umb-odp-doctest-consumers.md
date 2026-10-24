# 0362umb — ODP doctest consumer migration

The final ODP leaf gate at `f6fe430a4d` found eight stale documentation
consumers.  Seven examples in `authoring/mutable.rs` treated the fallible
`slides()` projection as a slice, and the `package/presentation.rs` example
named the private implementation alias `CoreResult` without importing it.

The examples now bind `transaction.slides()?` once before checking length,
emptiness, or indexing.  The presentation example uses the public
`std::result::Result<(), litchi_odp::Error>` spelling.  A scan of the other
ODP documentation found no remaining direct `slides()` consumers or
`CoreResult` references.  This is a documentation-consumer migration only;
no production behavior or approved API changed.

## Validation

The tested source is `b5207b27c1` (`fix(odp): migrate stale doctest
projections`), based on `410bd99f00`.  Commands used Rust 1.95.0, Cargo
1.95.0, the coordinator cargo shim first on `PATH`, one Cargo build job, one
test thread, `LITCHI_MEM_MAX=20G`, disabled incremental/debug information,
`--offline --locked`, and the separate target and temporary roots below.

* `cargo test --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-odp-doctests/target -p litchi-odp --all-features --doc`: 19 passed, 0 failed, exit 0, 43.74 seconds, maximum RSS 630,252 KB.
* `rustfmt --edition 2024 --check crates/litchi-odp/src/authoring/mutable.rs crates/litchi-odp/src/package/presentation.rs`: passed, exit 0.
* `cargo check --offline --locked --target-dir /home/zhuhe/code/litchi-worktrees/.coord/final-odp-doctests/target -p litchi-odp --all-features --all-targets`: passed, exit 0, 27.45 seconds, maximum RSS 501,524 KB.
* `git diff --check`: passed.

The ignored checkout lock was copied from the clean budget-validation input
`/home/zhuhe/code/litchi-worktrees/uo2-budget-validation/Cargo.lock`; it was
139,591 bytes with SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6` before
and after the checks, and remained byte-identical.  The external target was
410,690,676 bytes after validation.  The shim is
`/home/zhuhe/code/litchi-worktrees/.coord/bin/cargo` (520 bytes,
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`), and
the source `.cargo/config.toml` is 470 bytes with SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`.

All raw logs remain outside Git under
`/home/zhuhe/code/litchi-worktrees/.coord/final-odp-doctests/receipts/`.
Their compact manifest is `receipt-manifest.log` (2,165 bytes,
SHA-256 `87bc6081b312c50e5c5b2a4912c0bd8769d75bbf52888461b2424be99f07fbf1`).
