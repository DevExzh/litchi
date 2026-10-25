# 0362ul89 — ZIP acceptance and combined PPTX source preparation

Date: 2026-10-25. ZIP86 is accepted on coordinator integration. PPTX89 is
private source preparation; its runtime, corpus and performance acceptance
remain open. Nothing is pushed.

## ZIP acceptance

Atomic proposal `04879dd5036d5d79b351671f4d435e3db7af8674` merges at
`76b2d275f9652118e6fe4dbaba839aef00402b69`. Root verifies its exact parent
`ad138ffcedb5afc8ddbdbcc69c1e0c27e064cc77`, its tree, all 22 exact tested
Rust/record blobs and both additive glossary changes. Every ZIP/ODF/OPC/ops
consumer moves with the removed compatibility aliases. The 69 glossary
profiles and three ZIP name bans remain intact.

All eight actual guarded rows pass on
`e49a1062ecf59a8cf458e564af85d8880c4e9adb`: scoped formatting, ZIP tests and
strict Clippy, ODF streaming tests, OPC source-backed tests, scoped ops/facade
checks and the ZIP glossary. There are 632 ZIP, 14 ODF and 302 OPC tests,
948 total. Peak/final target allocation reported by the runner is
835,851,100 bytes. Root independently verifies source and lock identities,
logs, traces, target guards and empty terminal process ownership. The 99
selected glossary/inventory/census/fixture-guard tool tests also pass.

The actual state remains outside Git at
`.coord/0362ul86-f5-actual-e49a1062ec` in the worktrees parent directory:

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| `plan.json` | 8,143 | `e0ccfdaee37127cb2a2e8fd2dc2bae440087b448813788d58ebe0d0fb3147704` |
| `manifest.json` | 4,582 | `8eb253ccb374053c6d2cce579f6acfccd4b661fb41f761ef489185746447753a` |
| `rows.json` | 30,562 | `31fd779c4dcf66cc8f49a8affb464d178e737382c2be2db7766e59d6e240668d` |
| `result.json` | 32,777 | `e61de4e9dc48a1a810b0c1ed05b17643a40e632fd2453f44581d0cc484da20f8` |

The earlier workspace-format failure is retained in 0362ul86c. Its 335-path
formatting debt remains required before the final gate; scoped formatting
does not waive it.

## PPTX composition

The coordinator worktree `uo2-pptx-combined-0362ul89` composes reviewed
61h (`6139650eea`), 62g (`790f694f88`), 63 (`0d38e80bd8`), structural 81c
(`bf41322310`) and history 82a (`fac5957f84`) over integration `ad138ffcedb`.
The intermediate source merge is `126647b3f4`; it is not an acceptance merge.

Conflict resolutions retain all 13 semantic profiles behind the shared Arc,
physical part-byte accounting and persistent replay indexes together,
recipient reservation/laser policy checks, and both shape and tag limits at
the single raw shape-tag entry. All affected example and raw test callers
migrate with that entry. The glossary retains 69 profiles and 77 required
signatures in each PPTX profile. All 99 selected tooling tests pass on this
composition; Cargo has not validated it yet.

Composition review identifies additional required fixes. `f53cd281ee`
retains the complete semantic bundle and cache policy across deferred and
SVG physical reopening, and makes metadata publication recompute physical
part bytes through the shared derivation path. Both metadata callers adopt
its fallible result. `d8b106b8ff` adds a regression for deferred streaming,
materialization, non-pristine SVG reopening, retained policies and a selected
shape limit that the default policy would admit. This regression is committed
but has not run.

The 89a/b return (`568633279d`, `ffc20062f9`) composes at `8a9b5438f4`.
Laser edits retain one affected-slide usage ledger, charge each successful
candidate once and keep failed/retried/reset edits atomic. Prepared slides
propagate selected-policy/MCE admission errors before reservation and copy,
while preserving enclosed whole-shape removal. Independent review accepts
the laser logic; root verifies the corrected admission order. Runtime and
performance validation remain required.

Typed error work 84 remains separate. Reviews return stale helper calls,
incorrect resource dimensions, fabricated finite-over-budget observations
and an unowned maximum sentinel. Repairs 84a/b require compiler validation
and further source review before composition. The historical cross-copy
transcript must be regenerated from the actual composed source, reviewed and
committed before its final acceptance capture.

## Remaining acceptance

Shared OOXML properties and Web/Ribbon pass scoped source review and are
entering actual validation. Properties preserves the failed formatter run
on `e13bd7e843` and compiler failure on `a3eff9e1a6`; the latter exposes a
stale private `clear` call, repaired at `27be5389f7`. Web/Ribbon's first
attempt is refused before row zero because another build is active. No code
failure or test pass is inferred from that scheduling rejection. These
isolated actual captures must run serially. VBA87c remains under required
recipient-limit and source-fidelity repair.

The complete current gate, isolated workspace documentation, 84-profile raw
census, separate 68-package inventory, corpus/native execution and quiet ABBA
measurements remain open. Only redaction audit 22 is deferred.
