# 0362ul89s — PPTX shape consumer diagnostics

Date: 2026-10-25. This is a required source-preparation repair, not runtime,
corpus or performance acceptance. Only redaction audit 22 is deferred.

The actual two-row F5 capture started from
`15d8d53a0cf6d27f350b866463b4235bf26b08d9` on
`uo2/pptx-combined-0362ul89`. Scoped formatting passed. The keep-going
`cargo check --keep-going --offline --locked -p litchi-pptx --all-targets
--all-features` row failed after reporting 130 diagnostics. The assigned
shape-consumer subset contained 45 diagnostics: removed Designer package
wrappers, unnecessary qualifications in the shape and tag tests, and the
new typed `observed` field in the laser limit pattern.

Commit `e7ade643faafd3277eabeb1945fd54abb7e7d64d` repairs only these five
source files:

- `crates/litchi-pptx/src/shape/designer/tests.rs` migrates package tests to
  `load_with_limits`, `put_with_limits` and `remove_with_limits`, preserving
  the selector and value assertions.
- `crates/litchi-pptx/src/shape/owned/tests.rs` and
  `crates/litchi-pptx/src/shape/mod.rs` use the in-scope `Limits::DEFAULT`
  spelling without changing the selected policy.
- `crates/litchi-pptx/src/laser/edit.rs` and
  `crates/litchi-pptx/src/tag/tests.rs` use the family-scoped typed limit
  kind; the laser replay assertion ignores only the newly required observed
  field, while the exact observed assertions remain intact.

The source commit has no runtime or performance evidence. A fresh combined
compiler/test capture, strict Clippy, corpus execution and final acceptance
remain required.

The actual capture remains outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-actual-15d8d53a0c`.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `diagnostics-index.json` | 18,586 | `c25bee0ed0c486cd54e2d132c928509bfb08484991437136c53c403104b5d3be` |
| `01-pptx-all-features-check-attempt-actual-f5-01.log` | 100,739 | `a31129804840cddda84361dc4201852856d55afc25de7f0d49cc7d5868295084` |
| `result.json` | 33,327 | `2bbc5513db14610cef06bebbeed636c8388c0cb2051d02570e4719fa774de266` |

The source worktree is clean at the follow-up commit. No Cargo, rustdoc or
build command was run for this repair.
