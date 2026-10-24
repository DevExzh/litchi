# 0362uvb — final format error extensibility

- Date: 2026-10-24.
- Base: `945eb1d6b9ef6c56fa26c82e12758d38c1b9015d`.
- Contract: 0351 §2.4 and API glossary §5.
- Status: the 18 assigned exhaustive-error rows are repaired; the fonts
  obfuscation row is a separate concurrent lane.

The public errors reported by the final API inventory now carry
`#[non_exhaustive]` in these crates: codepage; crypto (integrity, labels,
legacy RC4, protected and RC4); Keynote media playback; Numbers' seven cell
data-format errors and the Numbers wire error; PPT text editing; VBA; and
XLSX sort conditions. This is an API-extensibility boundary only: variants,
payloads, display text and validation behavior are unchanged.

The compiler identified every exhaustive downstream match. Numbers-wire
allocation, output-limit and parse mappings retain their existing categories
in Numbers and the IWA adapter. Legacy and CryptoAPI RC4 mappings in DOC, XLS
and PPT retain their current categories. A future variant falls back to the
existing broad category with a fixed message where the destination carries a
message; it never exposes authored content. The PPT text-edit classifier
continues to classify known refusals by `RefusalKind`, with an unknown future
variant classified as corrupt input.

For codepage, crypto integrity/labels/protected, Keynote playback, the seven
semantic Numbers errors, VBA and XLSX sort, source inspection found only
construction, propagation, or defining-crate handling and no external
exhaustive match. The all-target/all-feature compiler check below covers the
workspace crates and proves that no `crates/*` consumer was left without a
wildcard arm. The standalone fuzz manifests were audited separately: their
affected-error uses are non-exhaustive-safe predicates (the duration
`ReversedRange` checks are not exhaustive matches). No inventory row is
silently suppressed; the separate fonts row remains visible for its owner.

## Validation

The leaf was checked with the memory-capped, single-job environment and an
external target:

```text
env CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/exhaustive-error-own/target CARGO_BUILD_JOBS=1 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G PATH=/home/zhuhe/code/litchi-worktrees/.coord/bin:$PATH cargo check --workspace --all-targets --all-features --locked
```

It passed after the consumer migrations. `cargo fmt --all -- --check`,
`git diff --check`, the no-3rdparty-reference test and the API glossary unit
tests are run before commit. No fixture, compatibility alias, lint
allowance, performance-sensitive route or redaction rule changed.
