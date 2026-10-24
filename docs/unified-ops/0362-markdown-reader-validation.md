# 0362 — Markdown reader shared commit validation receipt

Commit `83343f2ff17848c81312d4cdd7116e36a39d7a81` repairs the three Markdown
reader inventory rows.  `reader::Commit` is now the shared
`litchi_core::edit::Commit<Snapshot, Patch, Diagnostics>` triad; the retired
`snapshot()` accessor is gone; history, replay, README examples, and all
normative/reader consumers use `package()`, `into_package()`, and
`into_parts()`.  Markdown's `changed`, `touched_blocks`, and
`full_reparse_performed` diagnostics remain intact.

## Source identity and evidence boundary

The package test suite ran on the working candidate before the commit was
created: base `945eb1d6b9ef6c56fa26c82e12758d38c1b9015d` plus the exact four-file
diff later committed as `83343f2ff1`.  No source edit occurred between the test
completion and that commit.  The source hash was not captured inside the test
log, so this is a historical source binding rather than an independently
rerun post-commit test.  The reconstructed external patch is
`/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/markdown-reader-945-to-833.diff`
(17,487 bytes, SHA-256
`2a76cf166934e38e4be4af58d40e40011eabdfe3ba659ab22f5c33f7b9bdab05`).

The rustdoc and inventory checks ran after the commit on the exact
`83343f2ff1` tree.  The coordinator's final composed gate remains the
authoritative post-merge validation.

## Commands and inputs

The test and rustdoc invocations used Rust 1.95.0 (`rustc` commit
`59807616e`, Cargo commit `f2d3ce0bd`), the cargo shim first on `PATH`, one
Cargo build job, one test thread, `LITCHI_MEM_MAX=20G`, disabled incremental and
debug information, and offline locked resolution.  The environment script is
`/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/env.sh` (489 bytes,
SHA-256
`3b2855875112ff5e830c55e82e22baec0e5711252023099d83a06a6954539008`), with
`CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/target`
and `TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/tmp`.
The shim is 520 bytes, SHA-256
`1b22a69fd625d57565041bf6c634d8ebd1a82be8b3d0006390472c86324230d8`; the
checkout `.cargo/config.toml` is 470 bytes, SHA-256
`50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d`.

The ignored lock was copied from
`/home/zhuhe/code/litchi-worktrees/uo2-budget-validation/Cargo.lock`; it is
139,591 bytes with SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
The lock remained byte-identical after the commands.  The environment script
does not set `LITCHI_BLESS`; because the historical shell did not explicitly
clear inherited variables, any future rerun must use `env -u LITCHI_BLESS`.

Exact commands (the environment script supplied the target and memory
settings) were:

```text
cargo test --offline --locked -p litchi-markdown --all-features
RUSTC_BOOTSTRAP=1 cargo rustdoc --offline --locked -p litchi-markdown --lib --all-features -- -Zunstable-options --output-format json --document-hidden-items
python3 tools/api_debt_inventory.py /home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/target/doc/litchi_markdown.json
```

## Results and external logs

The test command exited 0: 19 unit tests, 3 conformance tests, 3 normative
tests, 24 reader tests, and 21 doctests passed (70 total).  Its observed log
window was 08:40:20–08:40:35 UTC; no RSS sample was captured, so no memory or
performance claim is made.  Rustdoc exited 0 in the observed 08:40:59–08:41:10
UTC window.  The inventory exited 0 and emitted only its header, with no debt
rows.  The target measured 4,904,409,972 apparent bytes (4,923,977,728 bytes
on disk) after validation, below the 20 GB job bound.

Raw logs remain outside Git:

| command | receipt | bytes | SHA-256 |
| --- | --- | ---: | --- |
| package tests (pre-commit working candidate) | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/markdown-reader-tests-945eb1d6b9.log` | 7,208 | `d649b57bb40548ae4aedffeb99a1bcae413e66b23897dbfe7c87677d03ac9d3a` |
| rustdoc (exact `83343f2ff1`) | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/markdown-reader-rustdoc-83343f2ff1.log` | 789 | `d17ea758a3830f94d12cf5a2f8f21af861576065a9a764f52469594811dd7392` |
| inventory (exact `83343f2ff1`) | `/home/zhuhe/code/litchi-worktrees/.coord/docx-lazy-entry/markdown-reader-inventory-83343f2ff1.log` | 34 | `557372144b7879942bcfeb5ed3d4210bcd1d1657d42c879f1efd5bfae8f911c9` |

`rustfmt --edition 2021 --check` on the three changed Rust files, `git
diff --check`, and the repository size/binary guards also passed before the
commit.  This receipt records correctness evidence only; the full composed
gate must validate the merged source.
