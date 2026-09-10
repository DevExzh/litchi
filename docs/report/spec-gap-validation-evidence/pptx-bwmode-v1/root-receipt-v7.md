# PPTX p14:bwMode publication evidence v7

- Evidence directory: `/var/tmp/litchi-pptx-bwmode-publication-evidence-v7-20260910`
- Candidate worktree: `/var/tmp/litchi-pptx-bwmode-validation-20260910`
- Candidate base: `cf98bee37455e8e6d0e73002ff2e94299ef546cb`
- Root gate source manifest: `root-source-v7.json` (8 paths: 7 Rust source paths plus `FEATURE_MATRIX.md`)
- Root gate manifest: `root-gates-v7.json`
- No primary worktree edits, staging, or commits were made for this evidence bundle.

## Candidate

`bwmode-candidate-v7.patch.gz` is the gzip `-n -9` form of the candidate patch.
The uncompressed patch SHA-256 is
`6a7aeab3762b4f8fbbe53651c3009dc29beb52ef71cfc51cb3ded20dd83a26ef`.
The compressed artifact SHA-256 is
`8a73d682af1119a1a4b809167b9ac83b00188e762b23788a634c6f6c53e39710`.
`source-v7.tar.gz` contains the eight source-manifest paths and is reproducibly
metadata-normalized; its SHA-256 is
`8ecd5318b365fcfdf6fb1b8dfd188355976f118791390d2fb9070ce286d89b1d`.

The candidate provides typed inert `p14:bwMode` metadata for `p:contentPart`,
all 11 schema tokens, bounded canonical MCE `Requires="p14"` selection, source
preservation, inverse edits, and signature-policy behavior. It rejects typed
interpretation of unqualified/foreign lookalikes and keeps `p14:media` opaque.

## Root gates

The root v7 gates used Rust `1.95.0` with these exact commands:

```text
cargo +1.95.0 test -p litchi-pptx --all-features --no-fail-fast
cargo +1.95.0 clippy -p litchi-pptx --all-features --all-targets -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo +1.95.0 doc -p litchi-pptx --all-features --no-deps
```

Environment recorded in `root-gates-v7.json`:

```text
RUSTUP_HOME=/tmp/litchi-spec-gap-rustup
TMPDIR=/var/tmp/litchi-spec-gap-gate-tmp
CARGO_TARGET_DIR=/var/tmp/litchi-pptx-bwmode-root-target-195
CARGO_BUILD_JOBS=4
affinity=8-31
```

Root reported 877 tests and 6 doctests passing, with 0 failures and 2 ignored
doctests. The root test, Clippy, and rustdoc logs are compressed under `logs/`.
Their raw SHA-256 values are recorded in `root-gates-v7.json`; their compressed
SHA-256 values are recorded in `ARTIFACT_HASHES.tsv`.

## Local specification and review boundary

`spec-scope/SPEC_SCOPE.md` maps the copied local Markdown specifications to the
p14 owner, 11-token type, no-default declaration, media-owner distinction, and
native MCE example. `review-pending.md` explicitly records the unrelated
independent `xlsx_current_review` v7 verdict as **CLEAR with no findings**.
