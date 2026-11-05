# 0362ul89z — PPTX metadata strict-Clippy source repair

Date: 2026-10-25

## Scope

This source-only repair addresses the 26 diagnostics reported for
`crates/litchi-pptx/src/presentation_properties/metadata/` by the retained
0362ul89 qualification-v2 strict-Clippy capture. The capture was run against
source commit
`02b2489eada2bfcd3d1b9bc285b0ea0d3c7ae0d6` (tree
`05942a20a7599b39d23b407831dac81b600a61cd`). It reported 76 diagnostics for
the package and failed the strict-Clippy row; this record does not claim that
the package or the 0362ul89 candidate is accepted.

The retained external evidence is:

- state: `(deleted-coord-artifact)`
- command log: `02-pptx-strict-clippy-all-features-attempt-qualification-f5-02.log`, 59,454 bytes, SHA-256 `bfe533c51350b5e9bfdeeb60ec6db7485f033679a4962e47d0264b14751785f6`
- diagnostic index: `clippy-diagnostics-index.json`, 13,168 bytes, SHA-256 `8c08ff22b3d5942f141cc6e488944838ebf35aef6c6a5c7e476c6930c80ca6ce`
- command: `cargo clippy --offline --locked -p litchi-pptx --all-features --all-targets -- -D warnings`

## Repair

The source repair is commit `b8cfe14138ff5610a37243247473c42c327bed0d`
(`fix(pptx): clear metadata clippy diagnostics`) on the sparse branch
`uo2/pptx-metadata-clippy-89z`. The commit removes the redundant function
`#[must_use]` attributes, removes conversions around the already-`usize`
`SourceLimits::max_xml_attribute_bytes()` value, fixes the custom-show
conditional and the two guide needless borrows, and keeps the returned types'
existing `#[must_use]` contracts.

The `guides` and `sections` DOM `Node` values no longer copy an operation
`Limits` value into every node. Their parse entry points and retained opaque
node serializers now receive the same explicitly selected operation profile.
This keeps the `Content::Node(Node)` variants inline, adds no `Box` or other
per-node allocation, leaves vector capacities and source-fidelity serialization
unchanged, and has no default-profile fallback. The selected limits still
bound parsing and retained-node serialization at the same call sites.

An external source patch against the exact qualification base is retained at:

`(deleted-coord-artifact)/metadata-clippy-89z-b8cfe14138.patch`

It is 18,621 bytes and has SHA-256
`2a7b6dcf8bdd8ed5b72950d2c32b971edd0532ab88cb40de072e12722270b1d1`.

## Checks and qualification status

No Cargo command, build, rustdoc, test, benchmark, or Clippy run was made for
this repair. Source-only checks passed for all 17 changed Rust files with
`rustfmt --edition 2024 --config skip_children=true --emit stdout`, followed
by `git diff --check`.

The source commit and patch are ready for coordinator composition. A fresh
strict-Clippy run on the composed candidate is required before any acceptance
decision.
