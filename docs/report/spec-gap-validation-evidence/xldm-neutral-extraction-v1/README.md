# Portable XLDM extraction evidence (v4)

This directory is a reviewable, source-only evidence bundle. It was created
under `/var/tmp` from the frozen candidate at
`/var/tmp/litchi-xldm-extraction-validation-20260910`. No primary worktree, candidate source, staging area, or commit was
modified while assembling it.

The candidate base is `906be17c2408e863bb26f2faf8f79361809597e6`. The corrected v4 root manifest binds 46 paths:
28 present candidate files and 18 files deleted from the old XLSX XLDM tree.
`v4/source-manifest-full46.json` is the exact root manifest. The independent
`v4/source-hashes-full46.txt` records candidate hashes for present files and
base hashes for deleted files; its path count and hashes were checked while
building this bundle.

The patch carries tracked edits and new-file diffs. The deterministic source
archive carries the 21 untracked source files; its 21 paths match
`v4/untracked-source-sha256.txt`. A reviewer can apply the patch at the base
commit and restore the archive contents to reconstruct the candidate.

## v4 gate commands and environment

These are the exact v4 root commands recorded for the changed XLSX facade and
its regression test:

```
cargo +1.95.0 test -p litchi-xlsx --all-features --test xldm_native_storage
cargo +1.95.0 clippy -p litchi-xldm -p litchi-xlsx --all-features --all-targets -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo +1.95.0 doc -p litchi-xldm -p litchi-xlsx --all-features --no-deps
```

The recorded gate environment was CPU range `8-31`,
`RUSTUP_HOME=/tmp/litchi-spec-gap-rustup`,
`TMPDIR=/var/tmp/litchi-spec-gap-gate-tmp`,
`CARGO_TARGET_DIR=/var/tmp/litchi-xldm-extraction-root-target-195`, and
`CARGO_BUILD_JOBS=4`. No `--offline` or `--locked` flag is asserted for these
v4 commands.

The v4 API test passed 3 tests with zero failures or ignores. Clippy and
rustdoc passed with the flags shown above. The root receipt also records
Rust 1.95 formatting and git-diff checks passing, and limits the v4 source
delta to `crates/litchi-xlsx/src/package.rs` and
`crates/litchi-xlsx/tests/xldm_native_storage.rs`.

The v3 baseline receipt is preserved separately under `baseline-v3/`. Its
full log records 1,438 unit/integration tests plus 2 doctests, all passing
with zero failures or ignores, and its strict Clippy log is retained. The
exact v3 command line was not retained in a receipt, so this bundle does not
invent one.

## Review conclusion

The bounded XLDM extraction review is clear for this candidate. The v4 facade
adapter preserves the historical XLSX error type for the public path
classifier, and the new regression test exercises that boundary. The neutral
XLDM storage/model implementation and native storage behavior are covered by
the retained v3 full gate and source manifest. No additional XLSB-model
blocker was found in this evidence pass.

## Artifact hashes

- v4 patch: `f5c2acf9099a708aeb14f42845b72809f388eac01b4c0f5798a3d4f18e4b6a9a`
- v4 untracked source archive: `92cddd61ddcf686d559713e2a91535ad5f3031fec533c6bd427ca14da1b5c67e`
- corrected 46-path root manifest: `510ed7d0bfcf118c8adfd295180e82eb15467c13a045517f955a0c04ae5c164a`
- v4 gate receipt: `f7bc33713d0671eed615d48606263c017bf9521842cf4fe269a2556389fed28e`
- full 46-path text binding: `a33ee2ac934954731bcc88b01830d398af86e174f7ed9149208ae17e949a1962`
- full 46-path JSON binding: `aa0c9966205fdde3b33f8544966d5b0b24fdf6273e5f0815f52edb9b02f8cc4e`

`SHA256SUMS.txt` hashes every other file in this bundle.

Raw logs and the source patch are stored with a `.gz` suffix. Decompression
reproduces the original bytes and the raw hashes recorded in the receipts.
