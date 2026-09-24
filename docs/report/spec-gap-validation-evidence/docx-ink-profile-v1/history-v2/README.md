# DOCX Ink profile publication evidence v2

This is a portable evidence bundle for the frozen candidate at
`/var/tmp/litchi-docx-ink-profile-validation-20260910`. It contains the exact
root v1 and v2 gate receipts, their source manifests, deterministic compressed
copies of the gate logs, and the complete eight-path candidate patch. It does
not modify the primary checkout, the frozen candidate, or git staging.

Independent `calc_chain_review` remains pending. The v2 terminal gates recorded
here are validation evidence and do not represent that independent review as
complete.

## v1 full proof retained separately

`v1-full-proof/` preserves the actual root v1 receipt, source manifest, and all
three root logs. The receipt's exact commands and environment are unchanged.
The test log contains 1,669 ordinary test passes and 77 doctest passes (76
`litchi_docx` plus 1 `litchi_drawingml`), with no failures; 32 tests are marked
ignored in the raw log. The separate directory keeps this proof distinct from
the v2 delta.

## v2 proof

`v2/root-gates-v2.json` records the exact root commands and reports 188 tests
passed, zero ignored, strict all-target/all-feature Clippy, warning-denied
rustdoc, and format/diff success. `v2/root-source-v2.json` verifies the eight
source blobs and their hashes. The compressed logs are byte-verifiable against
the original receipt hashes using `LOG-HASHES.txt` after decompression.

The exact commands recorded by the root v2 receipt are:

```text
cargo +1.95.0 test -p litchi-docx -p litchi-drawingml --all-features ink --no-fail-fast
cargo +1.95.0 clippy -p litchi-docx -p litchi-drawingml --all-features --all-targets -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo +1.95.0 doc -p litchi-docx -p litchi-drawingml --all-features --no-deps
```

The receipt environment is pinned to Rust 1.95.0 with CPU affinity 8-31,
`RUSTUP_HOME=/tmp/litchi-spec-gap-rustup`, `TMPDIR=/var/tmp/litchi-spec-gap-gate-tmp`,
`CARGO_TARGET_DIR=/var/tmp/litchi-docx-ink-profile-root-target-195`, and four
Cargo build jobs.

## Patch

`v2/patch-v2.patch.gz` is a `git diff --binary` patch from
`04fdd02a25029ea6d5729487207690a4740d3a70` covering exactly the eight paths
listed in `v2/patch-manifest.json`. The v2 source manifest is copied verbatim as
`v2/root-source-v2.json`; `v2/source-manifest.json` is retained as a convenient
manifest-named copy.

## Normative scope

`NORMATIVE-SCOPE.md` records the local `[MS-ODRAWXML]` Markdown locations used
for brush units/defaults and the v1 EMMA projection boundary. The implementation
accepts `m`, `cm`, `mm`, `in`, `pt`, `pc`, `em`, and `ex`; `px` selects the
specified default. The v2 semantic path applies XML Schema whitespace collapse
to decimal, integer, and boolean values while retaining source XML exactly.
