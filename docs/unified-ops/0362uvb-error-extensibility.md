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
env CARGO_TARGET_DIR=(deleted-coord-artifact) CARGO_BUILD_JOBS=1 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 CARGO_INCREMENTAL=0 CARGO_NET_OFFLINE=true LITCHI_MEM_MAX=20G PATH=(deleted-coord-artifact):$PATH cargo check --workspace --all-targets --all-features --locked
```

It passed after the consumer migrations. `cargo fmt --all -- --check`,
`git diff --check`, and the three Python modules
(`tools.test_check_api_glossary`, `tools.test_api_debt_inventory` and
`tools.test_no_3rdparty_references`) also passed; the Python command ran 58
tests with no failures.

The six-package library test command used for this leaf completed without a
failure. Its externally built test executables were rerun directly (without
Cargo, so no second workspace build) to make the package results explicit:

| Package | Passed | Failed | Ignored |
| --- | ---: | ---: | ---: |
| `litchi-doc` | 1,232 | 0 | 2 |
| `litchi-iwa` | 1,697 | 0 | 0 |
| `litchi-numbers` | 490 | 0 | 4 |
| `litchi-ppt` | 1,256 | 0 | 1 |
| `litchi-xls` | 1,327 | 0 | 0 |
| `litchi-ops` | 1,005 | 0 | 0 |
| **Total** | **7,007** | **0** | **7** |

The ignored, untracked diagnostic `Cargo.lock` used for that check is 139,592
bytes with SHA-256
`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`.
It was generated offline for the diagnostic source at this leaf; it is not
committed. The final composed check remains pending with the frozen 139,591
byte lock SHA-256
`52ed0cacac15b5e440349bfc0f4aca2665bcf50d59d4321624e596bb410783c6`.
No fixture, compatibility alias, lint allowance, performance-sensitive route
or redaction rule changed.
