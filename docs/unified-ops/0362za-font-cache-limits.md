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

For the original pre-commit run, no command-output log files were redirected
or retained; its successful output was present only in the interactive run
transcript. Consequently that run has no claimed log size or log SHA-256.

## Pinned-lock rerun receipt

The generated `6d7895e6…6747da` lock was retained as a diagnostic only: it
differs from the frozen `52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`
lock in six registry versions (`cc`, `crypto-common`, `ctutils`,
`generic-array`, `pulldown-cmark` and `smallvec`). The rerun copied
`Cargo.lock` from `uo2-budget-validation` without regeneration and used the
pinned hash above (139,591 bytes).

At source `dbe846e1fc42691821e21bdd14ecd86b0b1457b0`, with a clean worktree,
the pinned-lock rerun passed all 216 focused tests: fonts 25, cache 10 and
archive library 181. It also passed the strict checks below, all with
`--offline --locked`, one Cargo job, the same shim/memory/target/TMPDIR
settings recorded above:

```text
cargo clippy --offline --locked -p litchi-fonts -p litchi-iwa-cache -p litchi-iwa-archive --all-targets --all-features -- -D warnings
RUSTDOCFLAGS=-Dwarnings cargo doc --offline --locked --no-deps --all-features -p litchi-fonts -p litchi-iwa-cache -p litchi-iwa-archive
cargo fmt --all --check
```

The external retained logs and per-command resource receipts are under
`/home/zhuhe/code/litchi-worktrees/.coord/font-cache-limits/validation-52ed/`.
The raw `metadata.after.txt` remains unchanged at 3,043 bytes, SHA-256
`31319dc48ed32444af2e9513d3d0b8841fdec357d016bfb90214bf3c61750794`. Its
intermediate self-entry reports the earlier 2,706-byte state and hash
`a6ada7d0447acb8a8f2a5fa78e55a392604be031db0901f18ae641191d247054`; that
self-entry was written before the remaining entries were appended and is not
the final binding. The authoritative separate manifest is
`metadata.final-manifest.txt`, which excludes itself to avoid self-reference:
2,954 bytes, SHA-256
`cc120b976a98b40526ad5f16c96d36025fabb3a3b0b146f092d74433fe8b49f6`.
It binds the immutable logs, receipts and both before/after metadata files.
The corrected pre-target-size metadata is retained as
`metadata.before.corrected.txt` (220 bytes, SHA-256
`843214e4c5810d8650c987c8e3c90b0c455b2ab3b98d5dd9682604823c3efcf8`); the
earlier 684-byte `metadata.before.txt` is retained as a diagnostic because its
first target-size command had a quoting typo.

The command receipt files are:

| Receipt | Bytes | SHA-256 |
| --- | ---: | --- |
| `01-fonts.receipt.txt` | 3,928 | `fb1fd745fdbed3576bc5f4cd40d621c7040dec417505b8c8f26605c4a91abd02` |
| `02-cache.receipt.txt` | 2,213 | `1ada8d38fea63d1be9d36f0da2bcff14f4bd0cac5c1be31f660e0829bd1cd9e4` |
| `03-archive.receipt.txt` | 17,361 | `dcafe3b8555ec4e902036f3b5479d5a44b3ddb8165f0cbd9b3c92149c803d63c` |
| `04-clippy.receipt.txt` | 1,990 | `ffa00d7ad957dccdbaedc6f4fc301026a66b69e314dc2abe093172d22399d421` |
| `05-doc.receipt.txt` | 1,842 | `34b2e164cfef34de2a72b1d88cb44e347dabbec6ced089e17c3bbf51a5611f40` |
| `06-fmt.receipt.txt` | 98 | `996e0642b577cab9fe7fd4ac58b37dc797605c4013a819e7046ca629db37be53` |

The target measured 1,071,572,749 bytes after this rerun, below the 20 GB
ceiling. This receipt supersedes the earlier pre-commit validation limitation;
the tests now cover the frozen dependency graph without claiming a post-merge
trunk run.
