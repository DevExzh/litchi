# 0362za — Explicit limits for font preparation and weighted cache

The final direct-library API audit found two non-substrate public
`*_with_limits` entries. Both now expose one canonical raw constructor/helper
with its limits required, as required by API contract 0351 §2.5:

| Before | After |
| --- | --- |
| `litchi_fonts::embedding::prepare_with(resolver, glyphs, mode)` and `prepare_with_limits(..., limits)` | `prepare_with(resolver, glyphs, mode, limits)` |
| `WeightedCache::new(max_weight)` and `new_with_limits(max_weight, max_flights)` | `new(max_weight, max_flights)` |

The default `prepare` system-font workflow and IWA archive cache pass their
previous defaults explicitly. Existing bounded callers retain their exact
profiles. The old names are removed without aliases or deprecations, and the
font obfuscation error is marked non-exhaustive so downstream matches retain a
typed future fallback. All font, cache, and archive tests were migrated in the
same change.

Validation:

- `cargo test --offline --locked -p litchi-fonts --all-features` — 25 passed.
- `cargo test --offline --locked -p litchi-iwa-cache` — 10 passed.
- `cargo test --offline --locked -p litchi-iwa-archive --lib` — 181 passed.

## Validation receipt

The test, Clippy and rustdoc commands ran before the implementation commit,
against baseline `945eb1d6b9ef6c56fa26c82e12758d38c1b9015d` plus the six source
files in this change. No source file changed between those commands and commit
`169a03dbf0d45cdb12ed56de86870d5d53e3aaf5`. The exact binary source diff
(`git diff --binary HEAD^ HEAD --` over those six files) hashes to
`806796f42aaa0368834808e2f2ef6452ebec86d98b109bef804b6cd0a145cd34`.
Therefore this is pre-commit patch validation; no test command was rerun on
the commit object itself.

The successful commands were:

```text
cargo test --offline --locked -p litchi-fonts --all-features
cargo test --offline --locked -p litchi-iwa-cache
cargo test --offline --locked -p litchi-iwa-archive --lib
cargo clippy --offline --locked -p litchi-fonts -p litchi-iwa-cache -p litchi-iwa-archive --all-targets --all-features -- -D warnings
RUSTDOCFLAGS=-Dwarnings cargo doc --offline --locked --no-deps --all-features -p litchi-fonts -p litchi-iwa-cache -p litchi-iwa-archive
cargo fmt --all --check
```

The first cache test attempt exposed only a test-helper type-inference error;
the helper was corrected and the command above then passed. The source test
artifacts were observed at `08:33:56Z` (fonts), `08:35:34Z` (cache) and
`08:37:08Z` (archive); these are artifact mtimes and bound the runs, not
claimed command-start timestamps. The target's latest artifact mtime was
`08:40:20Z`; a post-validation `du` measurement returned:

```text
du -sb /home/zhuhe/code/litchi-worktrees/.coord/font-cache-limits/target
984570413  .../target
```

The ignored lockfile was generated in this worktree with
`cargo generate-lockfile --offline` after the copied lock from the main
worktree was rejected by `--locked`; it is not committed. Its final SHA-256 is
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`, size
139,592 bytes. Commands used the shim-prefixed `PATH`, unset
`LITCHI_CAPPED`/`LITCHI_BLESS`, `LITCHI_MEM_MAX=20G`,
`CARGO_BUILD_JOBS=1`, `RUST_TEST_THREADS=1`, `CARGO_INCREMENTAL=0`, all three
Cargo debug levels set to `0`, `CARGO_NET_OFFLINE=true`, the external target
above and `TMPDIR=/home/zhuhe/code/litchi-worktrees/.coord/font-cache-limits/tmp`.
The toolchain was rustc/cargo 1.95.0 (`rustc` commit
`59807616e1fa2540724bfbac14d7976d7e4a3860`).

No command-output log files were redirected or retained; the successful output
was present only in the interactive run transcript. Consequently this record
does not claim a log size or log SHA-256.
